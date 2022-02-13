VERSION 5.00
Begin VB.Form Arena 
   BackColor       =   &H00008000&
   Caption         =   "Robot War Arena"
   ClientHeight    =   4872
   ClientLeft      =   2076
   ClientTop       =   432
   ClientWidth     =   5868
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4872
   ScaleWidth      =   5868
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   1560
   End
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

Dim myobject As Object
Dim colors(4) As Long

' values for scaling graphics. We don't use scale method
' (though it *would* be easier), 'cause we have a problem
' scaling and offsetting the shapes when the main window is
' resized. We like to suffer, too....
Dim xscale As Single
Dim yscale As Single
Dim yoffset As Single
Dim botoffset As Single
Dim bangoffset As Single

Dim Index As Integer

Private Type ShellStruct
    t As Integer
    dx As Single
    dy As Single
    tx As Single
    ty As Single
End Type

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
Private Sub Command3_Click()
    Dim dblStart As Double
    Dim dblEnd As Double
    Dim i As Long

dblStart = Timer        ' Get the start time.

For i = 0 To 9999
'        Arena.FillStyle = 0
'        Arena.FillColor = colors(1)
'        Arena.Circle (90, 90), 10, colors(1)
'        Arena.PaintPicture Pixmap(i).Picture, 90, 90
       Pixmap(1).Move 90, 90
       Pixmap(1).Move 90, 90
Next

dblEnd = Timer          ' Get the end time.

Debug.Print dblEnd - dblStart   ' Display the
                                    ' elapsed time.
End Sub
Private Sub Command2_Click()
    Dim dblStart As Double
    Dim dblEnd As Double
    Dim i As Long

dblStart = Timer        ' Get the start time.

For i = 0 To 9999
'        Arena.FillStyle = 0
'        Arena.FillColor = colors(1)
'        Arena.Circle (90, 90), 10, colors(1)
'        Arena.PaintPicture Pixmap(1).Picture, 90, 90
'       Pixmap(1).Move 90, 90
Next

dblEnd = Timer          ' Get the end time.

Debug.Print dblEnd - dblStart   ' Display the
                                    ' elapsed time.
End Sub
Private Sub Command1_Click()
    Dim dblStart As Double
    Dim dblEnd As Double
    Dim i As Long
    Dim r As String
Dim sdir As Single
Dim sres As Single
Dim arcstart As Single
Dim arcend As Single
Dim p1 As Single
Dim p2 As Single

dblStart = Timer        ' Get the start time.


For i = 0 To 9999

'Shape1.Move 90, 90

Next

dblEnd = Timer          ' Get the end time.

Debug.Print dblEnd - dblStart   ' Display the
Debug.Print p1, p2                                    ' elapsed time.
End Sub
    


Private Sub DrawScan(i As Integer, x As Integer, y As Integer)

Dim r As String
Dim sdir As Single
Dim sres As Single
Dim arcstart As Single
Dim arcend As Single
    
r = myobject.getscan(i)
    
    sdir = Val(Mid$(r, 1, 3))
    sres = Val(Mid$(r, 4, 2))
'    Arena.FillStyle = 1
    arcstart = ((sdir - sres + 360) Mod 360) / 57.3
    arcend = ((sdir + sres) Mod 360) / 57.3
'    Arena.Circle (x, y), 700, colors(i), -arcstart, -arcend

    Arena.ForeColor = colors(i)
    Arena.Line (x, y)-Step((Cos(arcstart) * 700), (Sin(arcstart) * 700))
    Arena.Line (x, y)-Step((Cos(arcend) * 700), (Sin(arcend) * 700))

End Sub
Private Sub DrawShell(i As Integer, t As Integer)

Dim r As String
Dim x As Single
Dim y As Single
Dim dx As Single
Dim dy As Single
Dim tx As Single
Dim ty As Single


End Sub


Private Sub Form_Load()

Dim i As Integer

colors(1) = RGB(255, 0, 0)
colors(2) = RGB(0, 255, 0)
colors(3) = RGB(0, 0, 255)
colors(4) = RGB(255, 0, 255)

Index = 0

Arena.Visible = True

Set myobject = CreateObject("RobotServer.Display")

'Arena.Scale (0, 999)-(999, 0)
xscale = Arena.ScaleWidth / 1000
yscale = Arena.ScaleHeight / -1000
yoffset = Arena.ScaleHeight

For i = 1 To 4
    Shell(i).Top = -100
    Pixmap(i).Top = -200
    bang(i).Top = -1000
    Pixmap(i).Height = Arena.ScaleHeight / 50
    Pixmap(i).Width = Arena.ScaleHeight / 50
    bang(i).Height = Arena.ScaleHeight / 12.5
    bang(i).Width = Arena.ScaleHeight / 12.5
Next

botoffset = Pixmap(1).Height / 2
bangoffset = bang(1).Height / 2
Debug.Print yoffset, yscale, botoffset

End Sub


Private Sub Form_Resize()

Dim i As Integer

'Arena.Scale (0, 999)-(999, 0)
xscale = Arena.ScaleWidth / 1000
yscale = Arena.ScaleHeight / -1000
yoffset = Arena.ScaleHeight

For i = 1 To 4
    Shell(i).Top = -100
    Pixmap(i).Top = -200
    bang(i).Top = -1000
    Pixmap(i).Height = Arena.ScaleHeight / 50
    Pixmap(i).Width = Arena.ScaleHeight / 50
    bang(i).Height = Arena.ScaleHeight / 12.5
    bang(i).Width = Arena.ScaleHeight / 12.5
Next

botoffset = Pixmap(1).Height / 2
bangoffset = bang(1).Height / 2

End Sub


Private Sub Form_Unload(Cancel As Integer)

Set myobject = Nothing

End Sub


Private Sub Timer1_Timer()

Dim r As String
Dim x As Single
Dim y As Single
Dim x2 As Single
Dim y2 As Single
Dim dir As Single
Dim fire As Integer
Dim scan As Integer
Dim i As Integer
Dim offset As Integer
Dim lastbot As Integer
Dim sdir As Single
Dim sres As Single
Dim arcstart As Single
Dim arcend As Single
Dim r2 As String

r = myobject.getarena

lastbot = Val(Mid$(r, 1, 1))

If (Mid$(r, 2, 1) = "R") Then
    offset = 2
'    Arena.Refresh
    For i = 1 To lastbot
        ' clear old scan, if any
        If Scans(i).x <> 0 Then
            Arena.ForeColor = Arena.BackColor
            Arena.Line (Scans(i).x, Scans(i).y)-Step(Scans(i).x2, Scans(i).y2)
            Arena.Line (Scans(i).x, Scans(i).y)-Step(Scans(i).x3, Scans(i).y3)
            Scans(i).x = 0
        End If
        ' process this frame
        If (Mid$(r, offset + 1, 1) <> "N") Then
            x = Val(Mid$(r, offset + 2, 3)) * xscale
            y = Val(Mid$(r, offset + 5, 3)) * yscale + yoffset
            fire = Val(Mid$(r, offset + 8, 2))
            scan = Val(Mid$(r, offset + 10, 1))
            Pixmap(i).Move x - botoffset, y - botoffset
            ' handle new scan this frame
            If scan > 0 Then
                r2 = myobject.getscan(i)
                sdir = Val(Mid$(r2, 1, 3))
                sres = Val(Mid$(r2, 4, 2))
                arcstart = ((sdir - sres + 360) Mod 360) / 57.3
                arcend = ((sdir + sres) Mod 360) / 57.3
                Arena.ForeColor = colors(i)
                ' x2 & y2 are offset values
                x2 = Cos(arcstart) * 700 * xscale
                y2 = Sin(arcstart) * 700 * yscale
                Scans(i).x = x
                Scans(i).y = y
                Scans(i).x2 = x2
                Scans(i).y2 = y2
                Arena.Line (x, y)-Step(x2, y2)
                ' x2 & y2 are offset values
                x2 = Cos(arcend) * 700 * xscale
                y2 = Sin(arcend) * 700 * yscale
                Arena.Line (x, y)-Step(x2, y2)
                Scans(i).x3 = x2
                Scans(i).y3 = y2
            End If
            ' handle new shot this frame
            If fire > 0 Then
                r2 = myobject.getshot(i)
                shells(i).tx = Val(Mid$(r2, 1, 3)) * xscale
                shells(i).ty = Val(Mid$(r2, 4, 3)) * yscale + yoffset
                shells(i).dx = Val(Mid$(r2, 7, 3)) * xscale
                shells(i).dy = Val(Mid$(r2, 10, 3)) * yscale
                shells(i).t = Val(Mid$(r2, 13, 2))
                ' Just for fun, show where it'll hit...
                Arena.Line (shells(i).tx - 100, shells(i).ty)-Step(200, 0), colors(i)
                Arena.Line (shells(i).tx, shells(i).ty - 100)-Step(0, 200), colors(i)
            End If
        End If
        offset = offset + 10
        
        ' Process any shells that are in flight.
        If shells(i).t > 0 Then
            fire = shells(i).t
            If fire <= 2 Then
                Shell(i).Move -250, 1000
                bang(i).Move shells(i).tx - bangoffset, shells(i).ty - bangoffset
'                Arena.Circle (shells(i).tx, shells(i).ty), 40, RGB(255, 0, 0)
            Else
                x = shells(i).tx - (shells(i).dx * (fire - 2))
                y = shells(i).ty - (shells(i).dy * (fire - 2))
                Shell(i).Move x, y
            End If
            fire = fire - 1
            If fire = 0 Then
                bang(i).Move -1000, 1000
            End If
            shells(i).t = fire
        End If
    
    Next i
End If

End Sub


