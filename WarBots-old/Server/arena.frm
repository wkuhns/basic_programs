VERSION 5.00
Begin VB.Form Arena 
   BackColor       =   &H00008000&
   Caption         =   "Robot War Arena"
   ClientHeight    =   4872
   ClientLeft      =   1548
   ClientTop       =   276
   ClientWidth     =   5868
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4872
   ScaleWidth      =   5868
   Begin VB.Shape bang 
      FillStyle       =   0  'Solid
      Height          =   371
      Index           =   4
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   374
   End
   Begin VB.Shape bang 
      FillStyle       =   0  'Solid
      Height          =   371
      Index           =   3
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   374
   End
   Begin VB.Shape bang 
      FillStyle       =   0  'Solid
      Height          =   371
      Index           =   2
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   374
   End
   Begin VB.Shape bang 
      FillStyle       =   0  'Solid
      Height          =   371
      Index           =   1
      Left            =   0
      Shape           =   3  'Circle
      Top             =   0
      Width           =   374
   End
   Begin VB.Shape Shell 
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   4
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   48
   End
   Begin VB.Shape Shell 
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   3
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   48
   End
   Begin VB.Shape Shell 
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   2
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   48
   End
   Begin VB.Shape Shell 
      FillStyle       =   0  'Solid
      Height          =   48
      Index           =   1
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   2280
      Width           =   48
   End
   Begin VB.Shape Pixmap 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   116
      Index           =   4
      Left            =   4680
      Shape           =   1  'Square
      Top             =   464
      Width           =   117
   End
   Begin VB.Shape Pixmap 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   116
      Index           =   3
      Left            =   4680
      Shape           =   1  'Square
      Top             =   464
      Width           =   117
   End
   Begin VB.Shape Pixmap 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   116
      Index           =   2
      Left            =   4680
      Shape           =   1  'Square
      Top             =   464
      Width           =   117
   End
   Begin VB.Shape Pixmap 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   116
      Index           =   1
      Left            =   4680
      Shape           =   1  'Square
      Top             =   464
      Width           =   117
   End
End
Attribute VB_Name = "Arena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' values for scaling graphics. We don't use scale method
' (though it *would* be easier), 'cause we have a problem
' scaling and offsetting the shapes when the main window is
' resized. We like to suffer, too....

Dim xscale As Single        ' multiplier for x coordinate
Dim yscale As Single
Dim yoffset As Single       ' y axis needs offset 'cause of title bar

Dim botoffset As Single     ' 1/2 width of a 'bot
Dim bangoffset As Single    ' 1/2 width of a bang

Dim Index As Integer

' Structure for holding data about an in-flight shell
Private Type ShellStruct
    t As Integer
    dx As Single
    dy As Single
    tx As Single
    ty As Single
End Type

' Structure for remembering a scan so we can erase it
Private Type ScanStruct
    x As Single
    y As Single
    x2 As Single
    y2 As Single
    x3 As Single
    y3 As Single
End Type

Dim shells(4) As ShellStruct

Dim Scans(4) As ScanStruct

Dim xrange As Integer
Dim yrange As Integer

Public Sub DrawMark(x As Integer, y As Integer, c As Long)
' Draw a 'plus' mark on the arena in color and coordinates given

If DebugState Then
    Arena.Line (x * xscale - 100, y * yscale + yoffset)-Step(200, 0), c
    Arena.Line (x * xscale, y * yscale + yoffset - 100)-Step(0, 200), c
End If

End Sub

Public Sub Form_Load()

Dim i As Integer

Index = 0

CalcScales

End Sub


Private Sub CalcScales()

Dim i As Integer

' We would *really* like to use this, but doesn't work...
' Arena.Scale (0, 999)-(999, 0)
' Instead, we do it by hand. Precalculate gain/offset values

xscale = Arena.ScaleWidth / 1000
yscale = Arena.ScaleHeight / -1000
yoffset = Arena.ScaleHeight

For i = 1 To 4
    ' Move 'em out of the way
    Shell(i).Top = -100
    Pixmap(i).Top = -200
    bang(i).Top = -1000
    ' Pixmaps are 'bots
    Pixmap(i).Height = Arena.ScaleHeight / 50
    Pixmap(i).Width = Arena.ScaleHeight / 50
    ' Bangs are explosion circles
    bang(i).Height = Arena.ScaleHeight / 12.5
    bang(i).Width = Arena.ScaleHeight / 12.5
Next

botoffset = Pixmap(1).Height / 2
bangoffset = bang(1).Height / 2

' range of cannon ( for scan arcs )
xrange = xscale * 700
yrange = yscale * 700

End Sub




Public Sub DrawFrame()
' Cyclic update routine for arena display

Dim x As Single
Dim y As Single
Dim x2 As Single
Dim y2 As Single
Dim dir As Single
Dim i As Integer
Dim arcstart As Single
Dim arcend As Single

    For i = 1 To LastBot
        ' clear old scan, if any, by drawing it in arena color
        If Scans(i).x <> -1 Then
            Arena.ForeColor = Arena.BackColor
            Arena.Line (Scans(i).x, Scans(i).y)-Step(Scans(i).x2, Scans(i).y2)
            Arena.Line (Scans(i).x, Scans(i).y)-Step(Scans(i).x3, Scans(i).y3)
            Scans(i).x = -1
        End If
        ' process this frame for each non-null 'bot
        If Bots(i).status <> "N" Then
            x = Bots(i).x * xscale
            y = Bots(i).y * yscale + yoffset
            Pixmap(i).Move x - botoffset, y - botoffset
            ' handle new scan this frame
            If Bots(i).scan > 0 Then
                Bots(i).scan = 0
                arcstart = ((Bots(i).sdir - Bots(i).sres + 360) Mod 360) / 57.3
                arcend = ((Bots(i).sdir + Bots(i).sres) Mod 360) / 57.3
                Arena.ForeColor = Bots(i).color
                ' x2 & y2 are offset values
                x2 = Cos(arcstart) * xrange
                y2 = Sin(arcstart) * yrange
                Scans(i).x = x
                Scans(i).y = y
                Scans(i).x2 = x2
                Scans(i).y2 = y2
                Arena.Line (x, y)-Step(x2, y2)
                ' x2 & y2 are offset values
                x2 = Cos(arcend) * xrange
                y2 = Sin(arcend) * yrange
                Arena.Line (x, y)-Step(x2, y2)
                Scans(i).x3 = x2
                Scans(i).y3 = y2
            End If
            ' handle new shot this frame
            If Bots(i).newshot > 0 Then
                Bots(i).newshot = 0
                shells(i).tx = Bots(i).tx * xscale
                shells(i).ty = Bots(i).ty * yscale + yoffset
                shells(i).dx = Bots(i).dx * xscale
                shells(i).dy = Bots(i).dy * yscale
                ' Just for fun, show where it'll hit...
                Arena.Line (shells(i).tx - 100, shells(i).ty)-Step(200, 0), Bots(i).color
                Arena.Line (shells(i).tx, shells(i).ty - 100)-Step(0, 200), Bots(i).color
            End If
        End If
        ' Process any shells that are in flight.
        If Bots(i).fire > 0 Then
            If Bots(i).fire > 2 Then
                x = shells(i).tx - (shells(i).dx * (Bots(i).fire - 2))
                y = shells(i).ty - (shells(i).dy * (Bots(i).fire - 2))
                Shell(i).Move x, y
            End If
            If Bots(i).fire = 2 Then
                Shell(i).Move -250, 1000
                bang(i).Move shells(i).tx - bangoffset, shells(i).ty - bangoffset
            End If
            If Bots(i).fire = 1 Then
                bang(i).Move -1000, 1000
            End If
        End If
    
    Next i

End Sub

Private Sub Form_Resize()

' User has resized us.

CalcScales

End Sub


