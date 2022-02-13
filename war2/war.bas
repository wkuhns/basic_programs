Attribute VB_Name = "Module1"
Option Explicit

Global combuff As String
Global Const INVERSE = 6
Global Const SOLIDMODE = 0

Type logpalette
    palversion As Integer
    palnumentries As Integer
    palpalentry(256) As Long ' Array length is arbitrary; may be changed
End Type

Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal h As Integer, ByVal start As Integer, ByVal entries As Integer, pal As Long) As Integer
Declare Function CreatePalette Lib "GDI32" (lpLogPalette As logpalette) As Long
Declare Function SetPaletteEntries Lib "GDI32" (ByVal hPalette As Integer, ByVal wStartIndex As Integer, ByVal wNumEntries As Integer, lpPaletteEntries As Long) As Long
Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Integer) As Integer

Global hmypal As Long
Global MyPalette As logpalette

Global pal(256) As Long
Global mypal(60) As Long

Type OrderStruct
    command     As String * 1
    n1          As Single
    n2          As Single
End Type

Type dlstruct                   ' display list
    Source      As String   '(S)ightings, (U)nit display, others as they come up
    unit        As Integer  ' -1 if not a unit
    X           As Integer
    Y           As Integer
    Shape       As Integer
    fgcolor     As Long
    bgcolor     As Long
    msg         As String
End Type

Type dcstruct                   ' display control
    used        As Integer      ' is this one active?
    mult        As Integer      ' does it represent multiple things?
    X           As Integer
    Y           As Integer
    Shape       As Integer
    fgcolor     As Long
    bgcolor     As Long
    msg         As String
End Type
    
Type typestruct
    name        As String
    speed       As Integer  ' how fast can it go
    air         As Integer  ' airborne?
    armor       As Integer  ' how much armor does it have
    wtype       As Integer  ' weapon type enumeration
    wrange      As Integer  ' weapon range
    wstrength   As Integer  ' weapon strength
    vision      As Integer  ' how far can it see
    buycost     As Integer  ' how much does it cost to buy
    usecost     As Integer  ' how much does it cost to use
    accuracy    As Integer  ' % how accurate is it
    range       As Integer  ' how far can it go
    fuelcap     As Integer  ' fuel capacity
    stealth     As Integer  ' % how easy to see?
    count       As Integer  ' How many do we have?
End Type

Type unitstruct
    type        As Integer  ' TYPE enumeration
    changed     As Integer  ' Did I change this tick?
    index       As Integer  ' index to this unit
    speed       As Single   ' current speed
    engaged     As Integer  ' am I in a battle?
    dx          As Single   ' destination
    dy          As Single
    health      As Integer  ' %
    fuel        As Integer  ' % fuel remaining
    side        As Integer  ' whose side?
    X           As Single   ' location
    Y           As Single
    camo        As Integer  ' am I camoflaged?
    attack      As Integer  ' how aggressive am I? (0-100)
    follow      As Integer  ' likeliness to follow enemy
    Retreat     As Integer  ' likeliness to retreat from battle
    ocount      As Integer  ' order count
    orders(10)  As OrderStruct
End Type

Type xypoint
    X           As Single
    Y           As Single
End Type

Type terrainstruct
    a           As Single      ' altitude
    t           As Integer      ' type
    d           As Integer      ' difficulty
End Type

Global Const axis = 100
Global rain(axis, axis)  As Integer
Global terrain(axis, axis) As terrainstruct

Global dlcount      As Integer
Global dl(100)      As dlstruct
Global dc(40)       As dcstruct

Global mapdx As Double
Global mapdy As Double
Global maptop As Integer
Global mapleft As Integer
Global MapWidth As Integer
Global MapHeight As Integer

Global LastClick    As xypoint

Global Const us = 0
Global Const THEM = 1

Global RED As Long
Global WHITE As Long
Global GREEN As Long
Global BLUE  As Long
Global CYAN  As Long
Global MAGENTA  As Long
Global YELLOW   As Long
Global BLACK    As Long
Global DKGREY   As Long
Global MEDGREY  As Long
Global LTGREY   As Long

Global Const DSQUARE = 1
Global Const DCIRCLE = 3

Global side As Integer

Global remote As Integer        ' 0 = no remote, 1 = master, 2 = slave
Global GlobalTime As Long

Global specs() As typestruct

Global Const asize = 50
Global army(2, asize) As unitstruct          ' army1
Global a1base   As xypoint
Global a2base   As xypoint
Global a1port   As xypoint
Global a2port   As xypoint


Global UnitWaiting As Integer
Global GlobalCommand As Integer

Global v(8) As Single                   ' vectors

Global mc(10) As Long                   ' map colors
Public Function getcom() As String

    If InStr(combuff, Chr$(10)) Then
        getcom = Left$(combuff, InStr(combuff, Chr$(13)) - 1)
        combuff = Right$(combuff, Len(combuff) - InStr(combuff, Chr$(10)))
    Else
        getcom = ""
    End If

End Function
Public Sub SendCom(buffer As String)

    MapForm!Winsock.SendData (buffer + Chr$(13) + Chr$(10))

End Sub

Sub AddDisplayItem(Source As String, X As Integer, Y As Integer, Shape As Integer, fgcolor As Long, bgcolor As Long, msg As String)

    Static i As Integer

    For i = 0 To dlcount - 1        ' Do we already have something here?
    If (Int(X) = Int(dl(i).X)) And (Int(Y) = Int(dl(i).Y)) Then
        dl(i).unit = -1             ' not a unit
        dl(i).Source = Source
        dl(i).Shape = Shape
        dl(i).fgcolor = fgcolor
        dl(i).bgcolor = bgcolor
        dl(i).msg = msg
        Exit Sub
    End If
    Next i
    
    If dlcount < 100 Then
        dl(dlcount).Source = Source
        dl(dlcount).unit = -1
        dl(dlcount).X = Int(X)
        dl(dlcount).Y = Int(Y)
        dl(dlcount).Shape = Shape
        dl(dlcount).fgcolor = fgcolor
        dl(dlcount).bgcolor = bgcolor
        dl(dlcount).msg = msg
        dlcount = dlcount + 1
    End If

End Sub

Sub AddDisplayUnit(unit As Integer, Shape As Integer, fgcolor As Long, bgcolor As Long)
    
    Static i As Integer

    For i = 0 To dlcount - 1        ' Do we already have the same unit?
        If dl(i).unit = unit Then
            dl(i).Shape = Shape
            dl(i).fgcolor = fgcolor
            dl(i).bgcolor = bgcolor
            Exit Sub
        End If
    Next i
    
    If dlcount < 100 Then
        dl(dlcount).Source = "U"
        dl(dlcount).unit = unit
        dl(dlcount).Shape = Shape
        dl(dlcount).fgcolor = fgcolor
        dl(dlcount).bgcolor = bgcolor
        If side = us Then
            dl(i).msg = Left$(specs(army(0, unit).type).name, 1)
        Else
            dl(i).msg = Left$(specs(army(1, unit).type).name, 1)
        End If
        dlcount = dlcount + 1
    End If

End Sub

Sub AddListItem(unit As unitstruct)
        
    Static ntext As String

    ntext = Format$(unit.index, "00 ")
    ntext = ntext + Format$(unit.side, "0 ")
    ntext = ntext + Format$(specs(unit.type).name, "!@@@@@@@@@ at ")
    ntext = ntext + Format$(Int(unit.X), "00 ")
    ntext = ntext + Format$(Int(unit.Y), "00")
    CommandForm!UnitBox.AddItem ntext

End Sub

Sub AddSightings()

    Static i As Integer
    Static X As Integer
    Static Y As Integer

    For i = 0 To MapForm!SightList.ListCount - 1
        X = Val(Mid$(MapForm!SightList.List(i), 24, 2))
        Y = Val(Mid$(MapForm!SightList.List(i), 28, 2))
'        x = Int(x * mapdx)
'        y = Int(y * mapdy)
        AddDisplayItem "S", X, Y, DSQUARE, RED, WHITE, Mid$(MapForm!SightList.List(i), 10, 1)
    Next i

End Sub
Sub DisplayMap2(index As Integer)
    
    Dim X As Integer
    Dim Y As Integer
    Dim h As Integer
    Dim ml As Single
    Dim mt As Single
    Dim dx As Single
    Dim dy As Single
    Dim yp As Single
    
'    ml = mapleft + mapdx / 3
'    mt = maptop + mapdy / 3
'    dx = mapdx / 4
'    dy = mapdy / 4
    
    ' Draw colors
    For Y = 0 To axis - 1
        yp = maptop + Y * mapdy
        For X = 0 To axis - 1
            h = Int(terrain(X, Y).d)     ' quantize
            If h < 10 Then
                MapForm.Line (mapleft + X * mapdx, yp)-Step(mapdx, mapdy), mc(h), BF
            End If
        Next X
        DoEvents
    Next Y
    DrawLines (1)
    DisplayRivers
    
End Sub

Sub DisplayMap()
    
    Dim X As Integer
    Dim Y As Integer
    Dim h As Integer
    Dim dkh As Integer
    Dim kw As Integer

    dkh = MapForm!Picture2.Height / 10
    kw = MapForm!Picture2.Width

    MapForm!Picture2.AutoRedraw = True
    MapForm.DrawWidth = 1
    
    ' draw map altitude key
    For Y = 0 To 9
        MapForm!Picture2.Line (0, Y * dkh)-Step(kw, dkh), mc(9 - Y), BF
    Next Y

    ' Draw colors
    For Y = 0 To axis - 1
        For X = 0 To axis - 1
            h = Int(terrain(X, Y).a / 10)     ' quantize
            If h > 0 And rain(X, Y) > 10 Then h = 0
            MapForm.Line (mapleft + X * mapdx, maptop + Y * mapdy)-Step(mapdx, mapdy), mc(h), BF
        Next X
        DoEvents
    Next Y
    DrawLines (0)
    DisplayRivers
End Sub

Public Sub DrawLines(index As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim yp As Single
    Dim xp As Single
    
    If index = 0 Then
        ' Draw horizontal lines
        MapForm.DrawWidth = 1
        For Y = 1 To axis - 1
            yp = maptop + Y * mapdy
            For X = 0 To axis - 1
                If Int(terrain(X, Y - 1).a / 10) <> Int(terrain(X, Y).a / 10) Then
                    MapForm.Line (mapleft + X * mapdx, yp)-Step(mapdx, 0), MEDGREY
                End If
            Next X
            DoEvents
        Next Y
    
        For X = 1 To axis - 1
            xp = mapleft + X * mapdx
            For Y = 0 To axis - 1
                If Int(terrain(X - 1, Y).a / 10) <> Int(terrain(X, Y).a / 10) Then
                    MapForm.Line (xp, maptop + Y * mapdy)-Step(0, mapdy), MEDGREY
                End If
            Next Y
            DoEvents
        Next X
    End If
    
    If index = 1 Then
        ' Draw horizontal lines
        MapForm.DrawWidth = 1
        For Y = 1 To axis - 1
            For X = 0 To axis - 1
                If Int(terrain(X, Y - 1).d) <> Int(terrain(X, Y).d) Then
                    MapForm.Line (mapleft + X * mapdx, maptop + Y * mapdy)-Step(mapdx, 0), MEDGREY
                End If
            Next X
            DoEvents
        Next Y
    
        For X = 1 To axis - 1
            For Y = 0 To axis - 1
                If Int(terrain(X - 1, Y).d) <> Int(terrain(X, Y).d) Then
                    MapForm.Line (mapleft + X * mapdx, maptop + Y * mapdy)-Step(0, mapdy), MEDGREY
                End If
            Next Y
            DoEvents
        Next X
    End If
    

End Sub

Sub DisplayRivers()
    Dim X As Integer
    Dim Y As Integer
    Dim h As Integer
    Dim dkh As Integer
    Dim kw As Integer

    'MapForm.DrawWidth = 2
    ' Draw colors
    For Y = 1 To axis - 2
        For X = 0 To axis - 2
            'h = 0
            ' No lines in swamps
            If terrain(X, Y).d <= 10 Then
                GoTo nextx
            End If
            MapForm.DrawWidth = terrain(X, Y).d - 10
            If terrain(X + 1, Y).d > 10 Then
                MapForm.Line (mapleft + (X + 0.5) * mapdx, maptop + (Y + 0.5) * mapdy)-Step(mapdx, 0), mc(0)
                'h = h + 1
            End If
            If terrain(X, Y + 1).d > 10 Then
                MapForm.Line (mapleft + (X + 0.5) * mapdx, maptop + (Y + 0.5) * mapdy)-Step(0, mapdy), mc(0)
                'h = h + 1
            End If
            If terrain(X + 1, Y + 1).d > 10 Then
                MapForm.Line (mapleft + (X + 0.5) * mapdx, maptop + (Y + 0.5) * mapdy)-Step(mapdx, mapdy), mc(0)
                'h = h + 1
            End If
            If terrain(X + 1, Y - 1).d > 10 Then
                MapForm.Line (mapleft + (X + 0.5) * mapdx, maptop + (Y + 0.5) * mapdy)-Step(mapdx, -mapdy), mc(0)
                'h = h + 1
            End If
'            If h > 2 Then
'                MapForm.Line (mapleft + x * mapdx, maptop + y * mapdy)-Step(mapdx, mapdy), mc(0), BF
'            End If
nextx:
        Next X
        DoEvents
    Next Y
'    MapForm.AutoRedraw = False
'    MapForm.DrawWidth = 3
'    MapForm.Line (mapleft, maptop)-Step(mapdx * 40, mapdy * 40), BLACK
'    MapForm.Print "B"
'    MapForm.DrawWidth = 1
'    MapForm.CurrentX = mapleft + 50 * mapdx
'    MapForm.CurrentY = maptop + 50 * mapdy
'    'MapForm.FontSize = 5
'    MapForm.FillColor = WHITE
'    MapForm.Line (mapleft + 50 * mapdx - 20, maptop + 50 * mapdy)-Step(MapForm.TextWidth("D") + 20, MapForm.TextHeight("D")), , B
'    MapForm.CurrentX = mapleft + 50 * mapdx
'    MapForm.CurrentY = maptop + 50 * mapdy
'    MapForm.Print "D"
    
End Sub

Sub GetColors()
    
    Dim pcount As Long
    Dim X As Integer
    Dim Y As Integer
    
    Dim rs As Integer
    Dim gs As Integer
    Dim bs As Integer

    Dim rf As Integer
    Dim gf As Integer
    Dim BF As Integer


    Dim pcolor As Long

    MyPalette.palversion = &H300
    MyPalette.palnumentries = 53

    rs = &H60
    gs = &HEA
    bs = &HEA

    rf = &H60
    gf = &H83
    BF = &HEA

    For X = 0 To 15
        pcolor = (rs + Int((rf - rs) / 15 * X)) * &H10000
        pcolor = pcolor + (gs + Int((gf - gs) / 15 * X)) * &H100
        pcolor = pcolor + (bs + Int((BF - bs) / 15 * X))
        MyPalette.palpalentry(X) = pcolor
        mypal(X) = pcolor
    Next X
    
    rs = &H60
    gs = &H83
    bs = &HEA

    rf = &H0
    gf = &H0
    BF = &HC0

    For X = 0 To 15
        pcolor = (rs + Int((rf - rs) / 15 * X)) * &H10000
        pcolor = pcolor + (gs + Int((gf - gs) / 15 * X)) * &H100
        pcolor = pcolor + (bs + Int((BF - bs) / 15 * X))
        MyPalette.palpalentry(X + 16) = pcolor
        mypal(X + 16) = pcolor
    Next X
    
    RED = RGB(255, 0, 0)
    WHITE = RGB(255, 255, 255)
    GREEN = RGB(0, 255, 0)
    BLUE = RGB(0, 0, 255)
    CYAN = RGB(0, 255, 255)
    MAGENTA = RGB(255, 0, 255)
    YELLOW = RGB(255, 255, 0)
    BLACK = RGB(0, 0, 0)
    DKGREY = RGB(63, 63, 63)
    MEDGREY = RGB(127, 127, 127)
    LTGREY = RGB(195, 195, 195)
    MyPalette.palpalentry(32) = RED
    MyPalette.palpalentry(33) = WHITE
    MyPalette.palpalentry(34) = GREEN
    MyPalette.palpalentry(35) = BLUE
    MyPalette.palpalentry(36) = CYAN
    MyPalette.palpalentry(37) = MAGENTA
    MyPalette.palpalentry(38) = YELLOW
    MyPalette.palpalentry(39) = BLACK
    MyPalette.palpalentry(40) = DKGREY
    MyPalette.palpalentry(41) = MEDGREY
    MyPalette.palpalentry(42) = LTGREY
    For X = 0 To 9
        MyPalette.palpalentry(43 + X) = mc(X)
    Next
            

    hmypal = CreatePalette(MyPalette)
    
'    Debug.Print "Status of CreatePalette: "; hmypal


'    pcount = SetPaletteEntries(hmypal, 0, 16, pal(0))
    
'    Debug.Print "Status of SetPaletteEntries: "; pcount

'    pcount = SelectPalette(ColorForm!Picture1.hDC, hmypal, True)
     pcount = SelectPalette(MapForm.hDC, hmypal, True)
'     pcount = SelectPalette(PickForm.hDC, hmypal, True)
'     pcount = SelectPalette(BuyForm.hDC, hmypal, True)
'     pcount = SelectPalette(CCCForm.hDC, hmypal, True)

    pcount = SelectPalette(MapForm!Picture2.hDC, hmypal, True)

'    Debug.Print "Status of SelectPalette: "; pcount

    pcount = RealizePalette(MapForm.hDC)

'    Debug.Print "Status of RealizePalette: "; pcount

'    pcount = GetSystemPaletteEntries(MapForm.hDC, 0, 255, pal(0))

'    Debug.Print "Status of GetSystemPaletteEntries: "; pcount

'   ColorForm.Show
'    ColorForm!Picture1.FillStyle = 0
'    ColorForm!Picture1.AutoRedraw = True
    For Y = 0 To 15
        For X = 0 To 15
            pcolor = pal(Y * 16 + X)
 '           Debug.Print Hex$(pcolor); " ";
'            ColorForm!Picture1.FillColor = pcolor
'            ColorForm!Picture1.Line (ColorForm!Picture1.Width * x / 16, ColorForm!Picture1.Height * y / 18)-Step(ColorForm!Picture1.Width / 16, ColorForm!Picture1.Height / 16), RGB(0, 0, 0), B
        Next X
'        Debug.Print
    Next Y

    For Y = 0 To 1
        For X = 0 To 15
            pcolor = mypal(Y * 16 + X)
'            Debug.Print Hex$(pcolor); " ";
'            ColorForm!Picture1.FillColor = pcolor
'            ColorForm!Picture1.Line (ColorForm!Picture1.Width * x / 16, ColorForm!Picture1.Height * (y + 16) / 18)-Step(ColorForm!Picture1.Width / 16, ColorForm!Picture1.Height / 16), RGB(0, 0, 0), B
        Next X
    Next Y


End Sub

Sub getgame()
    

End Sub


Sub InitGame()
    
    v(1) = 1 - axis
    v(2) = 0 - axis
    v(3) = -1 - axis
    v(4) = -1
    v(5) = 1
    v(6) = axis - 1
    v(7) = axis
    v(8) = axis + 1

    'mc(0) = &HEABA3A
    mc(0) = &HCC0000
    mc(1) = &H8FE43A
    mc(2) = &H3AE43A
    mc(3) = &H1ED74D
    mc(4) = &H1ED7A9
    mc(5) = &H1ED7D7
    mc(6) = &H60C8EA
    mc(7) = &H6083EA
    mc(8) = &H403CE3
    mc(9) = &HC0

    GetColors

    Randomize
    
    PauseTimer

    side = us

End Sub

Sub InitMap()
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim lowest As Integer
    Dim lx As Integer
    Dim ly As Integer
    Dim count As Integer
    Dim side As Integer
    Dim sum As Single

    For i = 0 To axis - 1
        For j = 0 To axis - 1
            terrain(i, j).a = 0
        Next j
    Next i
    
    ' make mountain ranges

    For i = 1 To 4 * Sqr(axis)          ' 30 ranges for 100x100

        ' offset up and to the left a little. Mountain ranges
        ' tend to grow down and to the right

        X = Rnd * axis - (Rnd * (axis / 10))
        Y = Rnd * axis - (Rnd * (axis / 10))
        If X < 0 Then X = Rnd * axis
        If Y < 0 Then Y = Rnd * axis

        terrain(X, Y).a = Rnd * 40 + 60
        For k = 1 To Rnd * axis * 4     ' 1.5 * axis points on avg
            j = Rnd * 100               ' 33% left, 37% right
            If j < 33 Then X = X - 1
            If j > 60 Then X = X + 1

            j = Rnd * 100               ' 33% up, 37% down
            If j < 33 Then Y = Y - 1
            If j > 60 Then Y = Y + 1

'           x = x + Int(Rnd * 3) - 1
'           y = y + Int(Rnd * 3) - 1

            If (X >= 0 And X < axis And Y >= 0 And Y < axis) Then
                terrain(X, Y).a = Rnd * 40 + 60
            End If
        Next k
    Next i

    ' make random lowlands
    For i = 1 To (axis * axis)
        X = Rnd * axis
        Y = Rnd * axis
        If terrain(X, Y).a = 0 Then
            terrain(X, Y).a = Rnd * 60
        End If
    Next i

    ' enlarge swamps
    
    For i = 0 To axis - 1
        For j = 0 To axis - 1
            If terrain(i, j).a = 0 Then
                If (i > 0) And (i < (axis - 1)) And (j > 0) And (j < (axis - 1)) Then
                    terrain(i - 1, j - 1).a = terrain(i - 1, j - 1).a * 0.9
                    terrain(i - 1, j).a = terrain(i - 1, j).a * 0.9
                    terrain(i, j - 1).a = terrain(i, j - 1).a * 0.9
                    terrain(i + 1, j - 1).a = terrain(i + 1, j - 1).a * 0.9
                    terrain(i + 1, j).a = terrain(i + 1, j).a * 0.9
                    terrain(i + 1, j + 1).a = terrain(i + 1, j + 1).a * 0.9
                    terrain(i, j + 1).a = terrain(i, j + 1).a * 0.9
                    terrain(i - 1, j + 1).a = terrain(i - 1, j + 1).a * 0.9
                End If
            End If
        Next j
    Next i
            

    ' make all spots average for surrounding terrain
    For i = 0 To axis - 1
        For j = 0 To axis - 1
            count = 0
            sum = 0
            For X = (i - 1) To (i + 1)
                For Y = (j - 1) To (j + 1)
                    If (X >= 0 And X < axis And Y >= 0 And Y < axis) Then
                        sum = sum + terrain(X, Y).a
                        count = count + 1
                    End If
                Next Y
            Next X
            terrain(i, j).a = sum / count
        Next j
    Next i
    
    ' Tilt towards left
    For i = 0 To axis - 1
        For j = 0 To axis - 1
            terrain(i, j).a = terrain(i, j).a + ((100 - terrain(i, j).a) / 100) * j * 20 / axis
        Next j
    Next i
    
    MakeRivers
    SetDifficulty
    sendmap

End Sub
Public Function max(a As Integer, b As Integer) As Integer
    If a > b Then
        max = a
    Else
        max = b
    End If
End Function

Public Function min(a As Integer, b As Integer) As Integer
    If a > b Then
        min = b
    Else
        min = a
    End If
End Function
Public Sub SetDifficulty()

    Dim i As Integer
    Dim j As Integer
    Dim delta As Integer
    Dim delta2 As Integer
    
    For i = 0 To axis - 2
        For j = 0 To axis - 2
            delta = (terrain(i, j).a - terrain(i + 1, j).a) / 2
            delta = max(delta, -9)
            delta = min(delta, 9)
            delta2 = (terrain(i, j).a - terrain(i, j + 1).a) / 2
            delta2 = max(delta2, -9)
            delta2 = min(delta2, 9)
            ' We are higher than cell to right
            If delta < -9 Or delta2 < -9 Then
                delta = delta
            End If
            If delta > 0 Then
                terrain(i, j).d = max(terrain(i, j).d, delta)
            Else
                terrain(i + 1, j).d = max(terrain(i + 1, j).d, -delta)
            End If
        
            If delta2 > 0 Then
                terrain(i, j).d = max(terrain(i, j).d, delta)
            Else
                terrain(i, j + 1).d = max(terrain(i, j + 1).d, -delta)
            End If
            If terrain(i, j).a <= 10 Then
                terrain(i, j).d = 10
            End If
        Next j
        If terrain(i, axis - 1).a <= 10 Then
            terrain(i, axis - 1).d = 10
        Else
            terrain(i, axis - 1).d = terrain(i, axis - 1).a / 10
        End If
    Next i
    For j = 0 To axis - 1
        If terrain(axis - 1, j).a <= 10 Then
            terrain(axis - 1, j).d = 10
        Else
            terrain(axis - 1, j).d = terrain(axis - 1, j).a / 10
        End If
    Next j

End Sub

Public Sub MakeRivers()
    Dim i As Integer
    Dim j As Integer

    MakeRain
    MakeFlow
    MakeFlood
    MakeRain
    MakeFlow
    MakeFlood
    MakeFlow
    MakeFlood
    ' Set difficulty for rivers.
    For i = 0 To axis - 1
        For j = 0 To axis - 1
            If terrain(i, j).d < 5 Then
                terrain(i, j).d = 1
                GoTo nextj
            End If
            If terrain(i, j).d < 20 Then
                terrain(i, j).d = 11
                GoTo nextj
            End If
            If terrain(i, j).d < 30 Then
                terrain(i, j).d = 12
                GoTo nextj
            End If
            If terrain(i, j).d < 50 Then
                terrain(i, j).d = 13
            Else
                terrain(i, j).d = 14
            End If
nextj:
        Next j
    Next i
    
End Sub

Public Sub MakeRain()
    Dim i As Integer
    Dim j As Integer
    
    ' make rain
    For i = 0 To axis - 1
        For j = 0 To axis - 1
            rain(i, j) = 1
        Next j
    Next i
    
    For i = 0 To axis - 1
        For j = 0 To axis - 1
            terrain(i, j).d = 0
        Next j
    Next i

End Sub

Public Sub MakeFlow()
    Dim i As Integer
    Dim j As Integer
    Dim k As Single
    Dim X As Integer
    Dim Y As Integer
    Dim z As Integer
    Dim lowest As Single
    Dim lx As Integer
    Dim ly As Integer
    Dim changes As Integer
    
'    Dim dx As Single
'    Dim dy As Single
'    Dim water As Single
'    Dim land As Single

For z = 1 To 40
    changes = 0
    ' Let water flow downhill
    For i = 1 To axis - 2
        For j = 1 To axis - 2
'            If i = 68 And j = 58 Then
'                i = i
'            End If
            ' find lowest surrounding cell
            If rain(i, j) = 0 Then
                GoTo nextj
            End If
            lowest = 9998
            k = terrain(i, j).a         ' save altitude
            terrain(i, j).a = 9999
            For X = i - 1 To i + 1
                For Y = j - 1 To j + 1
                    If terrain(X, Y).a < lowest Then
                        lowest = terrain(X, Y).a
                        lx = X
                        ly = Y
                    End If
                Next Y
            Next X
            ' flow rain downhill
            If lowest < k Then
                rain(lx, ly) = rain(lx, ly) + rain(i, j)
                If rain(lx, ly) > terrain(lx, ly).d Then
                    terrain(lx, ly).d = rain(lx, ly)
                End If
                rain(i, j) = 0
                changes = changes + 1
            End If
            terrain(i, j).a = k
nextj:
        Next j
    Next i
    If changes = 0 Then
        Exit Sub
    End If
Next z

End Sub

Public Sub MakeFlood()
    Dim i As Integer
    Dim j As Integer
    Dim k As Single
    Dim X As Integer
    Dim Y As Integer
    Dim z As Integer
    Dim x1 As Integer
    Dim x2 As Integer
    Dim y1 As Integer
    Dim y2 As Integer
    Dim lowest As Single
    Dim lx As Integer
    Dim ly As Integer
    Dim dx As Single
    Dim dy As Single
    Dim water As Single
    Dim land As Single

    ' Water has now flowed downhill. Let's look for spots where it's
    ' accumulated. If they're at altitude, we need to open a gully for them to flow
    ' downhill.
    For i = 0 To axis - 1
        For j = 0 To axis - 1
            If i = 67 And j = 58 Then
                i = i
            End If
            If rain(i, j) > 3 And terrain(i, j).a > 5 Then
                lowest = terrain(i, j).a
                For X = 1 To 10
                    For Y = 1 To 10
                        x1 = max(i - X, 0)
                        x2 = min(i + X, axis - 1)
                        y1 = max(j - Y, 0)
                        y2 = min(j + Y, axis - 1)
                        If terrain(x1, y1).a < lowest Then
                            lowest = terrain(x1, y1).a
                            lx = -X
                            ly = -Y
                        End If
                        If terrain(x2, y1).a < lowest Then
                            lowest = terrain(x2, y1).a
                            lx = X
                            ly = -Y
                        End If
                        If terrain(x1, y2).a < lowest Then
                            lowest = terrain(x1, y2).a
                            lx = -X
                            ly = Y
                        End If
                        If terrain(x2, y2).a < lowest Then
                            lowest = terrain(x2, y2).a
                            lx = X
                            ly = Y
                        End If
                        If terrain(x1, j).a < lowest Then
                            lowest = terrain(x1, j).a
                            lx = -X
                            ly = 0
                        End If
                        If terrain(x2, j).a < lowest Then
                            lowest = terrain(x2, j).a
                            lx = X
                            ly = 0
                        End If
                        If terrain(i, y1).a < lowest Then
                            lowest = terrain(i, y1).a
                            lx = 0
                            ly = -Y
                        End If
                        If terrain(i, y2).a < lowest Then
                            lowest = terrain(i, y2).a
                            lx = 0
                            ly = Y
                        End If
                        If lowest < terrain(i, j).a Then
                            Exit For
                        End If
                        If lowest < terrain(i, j).a Then
                            Exit For
                        End If
                    Next Y
                    If lowest < terrain(i, j).a Then
                        Exit For
                    End If
                Next X
                ' Should have closest point lower than us with offsets in lx, ly
                If lowest <> terrain(i, j).a Then
                    If Abs(lx) > Abs(ly) Then
                       k = Abs(lx)
                    Else
                        k = Abs(ly)
                    End If
                    dx = lx / k
                    dy = ly / k
                    ' erode a sloping channel from here to there
                    For z = 1 To k
                        X = max(Int(i + z * dx + 0.5), 0)
                        Y = max(Int(j + z * dy + 0.5), 0)
                        X = min(X, axis - 1)
                        Y = min(Y, axis - 1)
                        If (terrain(i, j).a - z * (terrain(i, j).a - lowest) / k) < lowest Then
                            X = X
                        End If
                        terrain(X, Y).a = terrain(i, j).a - z * (terrain(i, j).a - lowest) / k
                    Next z
                Else
                    ' no lower spot found - raise us up
                    terrain(i, j).a = terrain(i, j).a + 5
                End If
           End If
        Next j
    Next i
    
    
End Sub

Sub sendmap()
    Dim i As Integer
    Dim j As Integer

    For i = 0 To axis - 1
        For j = 0 To axis - 1
            SendCom "M" & Format(i, "000") & " " & Format(j, "000") & " " & Format(terrain(i, j).a, "0000")
        Next j
    Next i
    SendCom "D"
    
End Sub

Sub LoadGame()
    
    Dim Filename As String
    Dim i As Integer
    Dim j As Integer

    MapForm!FileDialog.Filter = "WarGame (*.war)|*.war"
    MapForm!FileDialog.Action = 1
    Filename = MapForm!FileDialog.Filename
    Open Filename For Input As #1

    ' terrain map
    For i = 0 To axis - 1
        For j = 0 To axis - 1
            Input #1, terrain(i, j).a
        Next j
    Next i

    Input #1, a1base.X, a1base.Y
    Input #1, a1port.X, a1port.Y

    ' Army 1
    For i = 0 To asize - 1
        Input #1, army(0, i).type, army(0, i).index
        Input #1, army(0, i).speed, army(0, i).engaged
        Input #1, army(0, i).dx, army(0, i).dy
        Input #1, army(0, i).health, army(0, i).fuel
        Input #1, army(0, i).side, army(0, i).X
        Input #1, army(0, i).Y, army(0, i).camo
        Input #1, army(0, i).attack, army(0, i).follow
        Input #1, army(0, i).Retreat, army(0, i).ocount
        For j = 0 To army(0, i).ocount - 1
            Input #1, army(0, i).orders(j).command, army(0, i).orders(j).n1, army(0, i).orders(j).n2
        Next j
    Next i

    Input #1, a2base.X, a2base.Y
    Input #1, a2port.X, a2port.Y

    ' Army 2
    For i = 0 To asize - 1
        Input #1, army(1, i).type, army(1, i).index
        Input #1, army(1, i).speed, army(1, i).engaged
        Input #1, army(1, i).dx, army(1, i).dy
        Input #1, army(1, i).health, army(1, i).fuel
        Input #1, army(1, i).side, army(1, i).X
        Input #1, army(1, i).Y, army(1, i).camo
        Input #1, army(1, i).attack, army(1, i).follow
        Input #1, army(1, i).Retreat, army(1, i).ocount
        For j = 0 To army(1, i).ocount - 1
            Input #1, army(1, i).orders(j).command, army(1, i).orders(j).n1, army(1, i).orders(j).n2
        Next j
    Next i

    Input #1, mapdx, mapdy
    Input #1, side

    For i = 0 To 7
        Input #1, v(i)
    Next i

    For i = 0 To 9
        Input #1, mc(i)
    Next i

    Close #1

End Sub

Sub LoadUnits()

    Dim unitfile As String
    Dim junk As String
    Dim i As Integer
    Dim ucount As Integer

    unitfile = "units.txt"

    On Error Resume Next
    Open unitfile For Input As #1
    
    If Err > 0 Then
        On Error GoTo 0
        MapForm!FileDialog.Filter = "Unit File (*.txt)|*.txt"
        MapForm!FileDialog.Action = 1
        unitfile = MapForm!FileDialog.Filename
        Open unitfile For Input As #1
    End If
    On Error GoTo 0
    
    Input #1, ucount
    ReDim specs(ucount)

    'MapForm!TypeMenu(0).Visible = True

    For i = 0 To ucount - 1
        Input #1, junk, specs(i).name
        Input #1, junk, specs(i).speed
        Input #1, junk, specs(i).air
        Input #1, junk, specs(i).armor
        Input #1, junk, specs(i).wrange
        Input #1, junk, specs(i).wstrength
        Input #1, junk, specs(i).vision
        Input #1, junk, specs(i).buycost
        Input #1, junk, specs(i).usecost
        Input #1, junk, specs(i).accuracy
        Input #1, junk, specs(i).range
        Input #1, junk, specs(i).fuelcap
        Input #1, junk, specs(i).stealth
        Input #1, junk
        'If i > 0 Then
        '    Load MapForm!TypeMenu(i)
        'End If
        
        'MapForm!TypeMenu(i).Caption = specs(i).name
        'MapForm!TypeMenu(i).Visible = True
        MapForm!DispBox(i).Text = Left$(specs(i).name, 1)
    
    Next i

    Close #1

End Sub
Public Sub UpdateDispBoxes()

    Dim i As Integer
    
  
    For i = 0 To UBound(specs)
        MapForm!DispBox(i).Text = Left$(specs(i).name, 1) & Format(specs(i).count, "00")
    Next i
    
End Sub
Sub MakePickList(index As Integer)

    Dim i As Integer
    Dim ucount As Integer

    CommandForm!UnitBox.Clear
    ucount = 0
    
    If side = us Then
        For i = 0 To asize - 1
            If (army(0, i).health > 0) And (army(0, i).type = index) Then
                AddListItem army(0, i)
                ucount = ucount + 1
            End If
        Next i
    Else
        For i = 0 To asize - 1
            If (army(1, i).health > 0) And (army(1, i).type = index) Then
                AddListItem army(1, i)
                ucount = ucount + 1
            End If
        Next i
    End If
    
    If ucount > 0 Then
        CommandForm!UnitBox.Selected(0) = True
        CommandForm.Show
    End If

End Sub


Sub PauseTimer()

    MapForm!Timer1.Enabled = False
    MapForm!RunButton.Enabled = True
    MapForm!PauseButton.Enabled = False

End Sub

Sub PlotDisplayItems()
    
    Static i As Integer
    Static j As Integer
    Static X As Single
    Static Y As Single
    Static pendcount As Integer
    Static pending(50) As dcstruct
    Static matched As Integer

    ' Map all items in global display list to pending control list. The
    ' pending control list has no duplicates for location.

    pendcount = 0
    For i = 0 To dlcount - 1
        matched = False
        ' if list item is a unit, set its current location
        If dl(i).unit <> -1 Then
            If side = us Then
                If army(0, dl(i).unit).health > 0 Then
                    dl(i).X = Int(army(0, dl(i).unit).X)
                    dl(i).Y = Int(army(0, dl(i).unit).Y)
                End If
            Else
                If army(1, dl(i).unit).health > 0 Then
                    dl(i).X = Int(army(1, dl(i).unit).X)
                    dl(i).Y = Int(army(1, dl(i).unit).Y)
                End If
            End If
        End If

        For j = 0 To pendcount - 1
            If (pending(j).X = dl(i).X) And (pending(j).Y = dl(i).Y) Then
                If dl(i).Source = "S" Then
                    pending(j).fgcolor = dl(i).fgcolor
                    pending(j).bgcolor = dl(i).bgcolor
                    pending(j).Shape = dl(i).Shape
                End If
                pending(j).msg = "*"
                pending(j).mult = True
                matched = True
                Exit For
            End If
        Next j

        ' no existing entry in pending list at this location. Add new entry.
        If Not matched And pendcount < 50 Then
            pending(pendcount).used = False     ' tracks allocation to controls
            pending(pendcount).mult = False
            pending(pendcount).X = dl(i).X
            pending(pendcount).Y = dl(i).Y
            pending(pendcount).Shape = dl(i).Shape
            pending(pendcount).fgcolor = dl(i).fgcolor
            pending(pendcount).bgcolor = dl(i).bgcolor
            pending(j).msg = dl(i).msg
            pendcount = pendcount + 1
        End If
    Next i

    ' Mark all controls as unused.

    For j = 0 To 39
        dc(j).used = False
    Next j

    ' For each pending item, see if there's a control already at the right spot
    ' If so, mark the pending item as done and the control as used.

    For i = 0 To pendcount - 1
        For j = 0 To 39
            If (pending(i).X = dc(j).X) And (pending(i).Y = dc(j).Y) Then
                dc(j).fgcolor = pending(i).fgcolor
                dc(j).bgcolor = pending(i).bgcolor
                dc(j).Shape = pending(i).Shape
                dc(j).msg = pending(i).msg
                pending(i).used = True
                dc(j).used = True
                Exit For
            End If
        Next j
    Next i

    ' For any pending units that have not been allocated to display controls,
    ' find unused controls and map pending unit data to them.

    j = 0
    For i = 0 To pendcount - 1
        If pending(i).used = False Then         ' this one not allocated yet
            While j < 39 And dc(j).used = True
                j = j + 1
            Wend
            If j < 40 Then
                dc(j).used = True
                dc(j).fgcolor = pending(i).fgcolor
                dc(j).bgcolor = pending(i).bgcolor
                dc(j).Shape = pending(i).Shape
                dc(j).X = pending(i).X
                dc(j).Y = pending(i).Y
                dc(j).msg = pending(i).msg
            Else
                WriteCCC "Display Control array full"
            End If
        End If
    Next i

    ' set display controls per list

    For i = 0 To 39
        If dc(i).used Then
            MapForm.DItem(i).Visible = True
            ' symbol will be 2 units on a side, so offset 1/2 unit
            X = dc(i).X * mapdx - (mapdx / 2) + mapleft
            Y = dc(i).Y * mapdy - (mapdy / 2) + maptop
            If (X <> MapForm!DItem(i).Left) Or (Y <> MapForm!DItem(i).Top) Then
                MapForm!DItem(i).Move X, Y
            End If
            MapForm!DItem(i).ForeColor = dc(i).fgcolor
            MapForm!DItem(i).FillColor = dc(i).bgcolor
            MapForm!DItem(i).Cls
            MapForm!DItem(i).CurrentX = (MapForm!DItem(i).Width - MapForm!DItem(i).TextWidth(dc(i).msg)) / 2 - 8
            MapForm!DItem(i).CurrentY = (MapForm!DItem(i).Height - MapForm!DItem(i).TextHeight(dc(i).msg)) / 2 - 5
            MapForm!DItem(i).Print dc(i).msg
        Else
            MapForm!DItem(i).Visible = False
        End If
    Next i

End Sub

Sub PrintMap()
    
    Dim X, Y As Integer
    Dim h As Integer
    Dim pdx As Single
    Dim mcolor As Long
    Dim b As Single

    pdx = Printer.Width / axis * 0.9
    If (Printer.Height / axis * 0.9) < pdx Then
        pdx = Printer.Height / axis * 0.9
    End If
    b = pdx * axis / 18

    For Y = 0 To axis - 1
        For X = 0 To axis - 1
'            h = 9 - (Int(t(x, y) / 10))     ' quantize
'            h = Int((h + 3) * 256 / 16)
'            mcolor = RGB(h, h, h)
'            Printer.Line (x * pdx + B, y * pdx + B)-Step(pdx, pdx), mcolor, BF
            h = Int(terrain(X, Y).a / 10)     ' quantize
            Printer.Line (X * pdx + b, Y * pdx + b)-Step(pdx, pdx), mc(h), BF
        Next X
        DoEvents
    Next Y

    For Y = 1 To axis - 1
        For X = 0 To axis - 1
            If Int(t(X, Y - 1) / 10) <> Int(t(X, Y) / 10) Then
                Printer.Line (X * pdx + b, Y * pdx + b)-Step(pdx, 0), 0
            End If
        Next X
        DoEvents
    Next Y

    For X = 1 To axis - 1
        For Y = 0 To axis - 1
            If Int(t(X - 1, Y) / 10) <> Int(t(X, Y) / 10) Then
                Printer.Line (X * pdx + b, Y * pdx + b)-Step(0, pdx), 0
            End If
        Next Y
        DoEvents
    Next X

    Printer.EndDoc
    WriteCCC "Print job spooled"

End Sub

Sub RemoveDisplayItem(i As Integer)

    Static j As Integer

    If (i >= 0) And (i < dlcount) Then
        j = i
        While (j < (dlcount - 2))
            dl(j) = dl(j + 1)
            j = j + 1
        Wend
    
        dlcount = dlcount - 1

    End If

End Sub


Sub SaveGame()

    Dim Filename As String
    Dim i, j As Integer

    PauseTimer

'    Open "map.dat" For Output As #1

    ' terrain map
'    For i = 0 To axis - 1
'        For j = 0 To axis - 1
'            Print #1, terrain(i, j); ", ";
'        Next j
'        Print #1, " "
'    Next i
    
'    Close #1

    MapForm!FileDialog.DefaultExt = ".war"
    MapForm!FileDialog.Filter = "WarGame (*.war)|*.war"
    MapForm!FileDialog.Action = 2
    Filename = MapForm!FileDialog.Filename
    Open Filename For Output As #1

    ' terrain map
    For i = 0 To axis - 1
        For j = 0 To axis - 1
            Print #1, terrain(i, j).a
        Next j
    Next i

    Print #1, a1base.X, a1base.Y
    Print #1, a1port.X, a1port.Y

    ' Army 1
    For i = 0 To asize - 1
        Print #1, army(0, i).type, army(0, i).index
        Print #1, army(0, i).speed, army(0, i).engaged
        Print #1, army(0, i).dx, army(0, i).dy
        Print #1, army(0, i).health, army(0, i).fuel
        Print #1, army(0, i).side, army(0, i).X
        Print #1, army(0, i).Y, army(0, i).camo
        Print #1, army(0, i).attack, army(0, i).follow
        Print #1, army(0, i).Retreat, army(0, i).ocount
        For j = 0 To army(0, i).ocount - 1
            Print #1, """; army(0,i).orders(j).command; """, army(0, i).orders(j).n1, army(0, i).orders(j).n2
        Next j
    Next i

    Print #1, a2base.X, a2base.Y
    Print #1, a2port.X, a2port.Y

    ' Army 2
    For i = 0 To asize - 1
        Print #1, army(1, i).type, army(1, i).index
        Print #1, army(1, i).speed, army(1, i).engaged
        Print #1, army(1, i).dx, army(1, i).dy
        Print #1, army(1, i).health, army(1, i).fuel
        Print #1, army(1, i).side, army(1, i).X
        Print #1, army(1, i).Y, army(1, i).camo
        Print #1, army(1, i).attack, army(1, i).follow
        Print #1, army(1, i).Retreat, army(1, i).ocount
        For j = 0 To army(1, i).ocount - 1
            Print #1, """; army(1,i).orders(j).command;""", army(1, i).orders(j).n1, army(1, i).orders(j).n2
        Next j
    Next i

    Print #1, mapdx, mapdy
    Print #1, side

    For i = 0 To 7
        Print #1, v(i)
    Next i

    For i = 0 To 9
        Print #1, mc(i)
    Next i

    Close #1

End Sub

Sub SendGame()
    
    Dim i As Integer

    For i = 0 To asize - 1
        SendUnitInfo army(0, i)
    Next i

    WriteCCC "All units sent"
    SendCom "DONE"

End Sub


Sub DispatchUnits(X As Single, Y As Single)

    Dim i As Integer
    Dim u As Integer
    Dim lx As Single
    Dim ly As Single
    
    lx = X - mapleft
    ly = Y - maptop

    If lx < 0 Then lx = 0
    If ly < 0 Then ly = 0

    lx = lx / mapdx
    ly = ly / mapdy

    If (lx > (axis - 1)) Then lx = axis - 1
    If (ly > (axis - 1)) Then ly = axis - 1

    For i = 0 To CommandForm!UnitBox.ListCount - 1
        If CommandForm!UnitBox.Selected(i) Then
            u = Val(Left$(CommandForm!UnitBox.List(i), 2))
            If side = us Then
                AddOrder army(0, u), "M", lx, ly
            Else
                AddOrder army(1, u), "M", lx, ly
            End If
 '           WriteCCC "Sent " + Str$(u) + " to " + Str$(Int(lx / mapdx)) + ", " + Str$(Int(ly / mapdy))
        End If
    Next i
    
End Sub

Sub SetupPlayer()
    
    MapForm!MessageBox.Text = "Click on main base"

    LastClick.X = 0

    While LastClick.X = 0
        DoEvents
    Wend

    MapForm!MessageBox.Text = ""

    AddDisplayItem "B", Int(LastClick.X), Int(LastClick.Y), DCIRCLE, BLUE, WHITE, "B"

    PlotDisplayItems

'    CommandForm.Show
'    CommandForm.UnitBox.Clear

    If side = us Then
        a1base = LastClick
        army(0, 0).type = 0                 ' general
        army(0, 0).health = 100
        army(0, 0).side = side
        army(0, 0).X = a1base.X
        army(0, 0).Y = a1base.Y
        army(0, 0).dx = a1base.X
        army(0, 0).dy = a1base.Y
        AddDisplayUnit 0, DSQUARE, BLACK, WHITE
        AddListItem army(0, 0)
    Else
        a2base = LastClick
        army(1, 0).type = 0                 ' general
        army(1, 0).health = 100
        army(1, 0).side = side
        army(1, 0).X = a2base.X
        army(1, 0).Y = a2base.Y
        army(1, 0).dx = a2base.X
        army(1, 0).dy = a2base.Y
        AddDisplayUnit 0, DSQUARE, BLACK, WHITE
        AddListItem army(1, 0)
    End If
'    CommandForm!UnitBox.Selected(0) = True

    PlotDisplayItems

End Sub

Sub StartTimer()

    MapForm!RunButton.Enabled = False
    MapForm!PauseButton.Enabled = True
    If remote = 2 Then
        While MapForm!PauseButton.Enabled = True
            tick
        Wend
    Else
        MapForm!Timer1.Enabled = True
    End If

End Sub

Sub tick()
    
    Static i As Integer
    Static msg As String
    Static unit As Integer
    
    Select Case remote
    Case 0                              ' No remote
        For i = 0 To asize - 1
            If army(0, i).health > 0 Then
                ' check for moving army 1 units
                If (army(0, i).X <> army(0, i).dx) Or (army(0, i).Y <> army(0, i).dy) Then
                    If MoveUnit(army(0, i)) Or army(0, i).engaged Then
                        ProcessMove army(0, i)
                    End If
                Else
                    If army(0, i).engaged Then
                        ProcessMove army(0, i)            ' do battle
                    End If
                End If
                ' Process pending commands
                If army(0, i).ocount > 0 Then
                    OrderProcess army(0, i)
                End If
            End If
        
            If army(1, i).health > 0 Then
                ' check for moving army 2 units
                If (army(1, i).X <> army(1, i).dx) Or (army(1, i).Y <> army(1, i).dy) Then
                    If MoveUnit(army(1, i)) Or army(1, i).engaged Then
                        ProcessMove army(1, i)
                    End If
                Else
                    If army(0, i).engaged Then
                        ProcessMove army(1, i)            ' do battle
                    End If
                End If
                ' Process pending commands
                If army(1, i).ocount > 0 Then
                    OrderProcess army(1, i)
                End If
            End If
        Next i
    
    Case 1, 2                              ' Remote
        
        For i = 0 To asize - 1
            If army(0, i).health > 0 Then
                ' check for moving army 1 units
                If (army(0, i).X <> army(0, i).dx) Or (army(0, i).Y <> army(0, i).dy) Then
                    If MoveUnit(army(0, i)) Or army(0, i).engaged Then
                        ProcessMove army(0, i)
                    End If
                Else
                    If army(0, i).engaged Then
                        ProcessMove army(0, i)            ' do battle
                    End If
                End If
                ' Process pending commands
                If army(0, i).ocount > 0 Then
                    OrderProcess army(0, i)
                End If
            End If
            If army(0, i).changed Then
                SendUnitInfo army(0, i)
                army(0, i).changed = False
            End If
            DoEvents
        Next i

    End Select

    ' Need to sync with remote.....
    PlotDisplayItems
    GlobalTime = GlobalTime + 1
    MapForm!ClockBox.Text = Format$(GlobalTime, "00\:00\:00")

End Sub


Sub WriteCCC(msg As String)
    
    If CCCForm!CCCBox.ListCount >= 25 Then
    CCCForm!CCCBox.RemoveItem 0
    End If

    CCCForm!CCCBox.AddItem msg

End Sub
Public Sub SwitchPlayers()
            
Dim i As Integer
            If side = us Then
                side = THEM
                MapForm!MessageBox.Text = "Them"
            Else
                side = us
                MapForm!MessageBox.Text = "Us"
            End If
            For i = 0 To asize - 1
                If army(0, i).health > 0 Then
'                    DrawCell Int(army(0,i).x), Int(army(0,i).y)
                End If
                If army(1, i).health > 0 Then
'                    DrawCell Int(army(1,i).x), Int(army(1,i).y)
                End If
            Next i

End Sub
