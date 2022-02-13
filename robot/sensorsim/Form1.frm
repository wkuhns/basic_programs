VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   13380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ssbtn 
      Caption         =   "single step"
      Height          =   375
      Left            =   10320
      TabIndex        =   3
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   7320
      Top             =   4800
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   5880
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   5880
      Width           =   3855
   End
   Begin VB.CommandButton ScanBtn 
      Caption         =   "Scan"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderWidth     =   6
      X1              =   3240
      X2              =   3240
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderWidth     =   6
      X1              =   5520
      X2              =   6120
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Shape Shape7 
      BorderWidth     =   5
      Height          =   1095
      Left            =   3000
      Top             =   1920
      Width           =   735
   End
   Begin VB.Shape Shape6 
      BorderWidth     =   5
      Height          =   615
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   4200
      Width           =   615
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   5
      Height          =   615
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   5
      Height          =   975
      Left            =   2520
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   5
      Height          =   975
      Left            =   6000
      Top             =   2280
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   5
      Height          =   2415
      Left            =   1080
      Top             =   1080
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   7
      Height          =   5295
      Left            =   360
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()

' Scale the arena to approximately 10x20 meters - units are cm.

xscale = Form1.ScaleWidth / 2000
yscale = Form1.ScaleHeight / -1000
yoffset = Form1.ScaleHeight

bot.x = 300
bot.y = 200
bot.t = 180
runflag = False
Timer1.Interval = 75
Timer1.Enabled = False

End Sub

Private Sub ScanBtn_Click()

If Timer1.Enabled Then
    Timer1.Enabled = False
Else
    Timer1.Enabled = True
End If

End Sub

Private Sub runme()

    doscan
    Form1.Refresh
    showbot
    planmove
    Text1 = "bot heading = " & bot.t
    movebot
    showscan
    DoEvents
    

End Sub

Private Sub movebot()

' move 5 cm per step

bot.x = bot.x + 5 * Sin(bot.t / 57.3)
bot.y = bot.y - 5 * Cos(bot.t / 57.3)

End Sub
Private Sub planmove()

Dim t As Integer
Dim l As Integer
Dim i As Integer
Dim s As Integer
Dim st As Integer
Dim lt As Integer
Dim xt As Integer

' First, are we clear of obstacles on either side?
' scan(17) and scan(18) are 5 degrees left and right of center
' scan(9) through scan(12) and scan(23) through scan(27) show
' what's approaching our left and right flanks
st = 9999

If scan(15) < 47 Then
    st = bot.t + 5
End If

If scan(20) < 47 Then
    st = bot.t - 5
End If

If scan(14) < 35 Then
    st = bot.t + 10
End If

If scan(21) < 35 Then
    st = bot.t - 10
End If

' How far to obstacle in front of us?

l = (scan(17) + scan(18)) / 2
lt = 9999

' is it farther to obstacle if we adjust our course?
' limit course changes to +/- 5 degrees

For i = 0 To 7
    If scan(17 + i) > l Then
        l = scan(17 + i)
        lt = bot.t + 5
    End If
    
    If scan(17 - i) > l Then
        l = scan(17 - i)
        lt = bot.t - 5
    End If

Next

xt = 9999

' now look for critical obstacles

If scan(13) < 28 Then
    xt = bot.t + 15
End If

If scan(22) < 28 Then
    xt = bot.t - 15
End If

If scan(12) < 24 Then
    xt = bot.t + 20
End If

If scan(23) < 24 Then
    xt = bot.t - 20
End If

If scan(11) < 23 Then
    xt = bot.t + 25
End If

If scan(24) < 23 Then
    xt = bot.t - 25
End If

If scan(10) < 23 Then
    xt = bot.t + 30
End If

If scan(25) < 23 Then
    xt = bot.t - 30
End If

Text2 = ""

If l < 40 Then
    Timer1.Enabled = False
    Text2 = "stop - long"
    For i = 0 To 35
        Debug.Print i * 10 - 175 & ": " & scan(i)
    Next
    xt = (lt + 180) Mod 360
End If

If lt <> 9999 Then
    bot.t = lt
    Text2 = "aim "
End If

If st <> 9999 Then
    bot.t = st
    Text2 = "adjust "
End If

If xt <> 9999 Then
    bot.t = xt
    Text2 = "avoid "
End If

    

For i = 0 To 35
    If scan(i) < 12 Then
        Timer1.Enabled = False
        Text2 = "stop - short"
        For l = 0 To 35
            Debug.Print l * 10 - 175 & ": " & scan(l)
        Next
        bot.t = (bot.t + -i + 10 - 5 + 180) Mod 360
    End If
    
Next


End Sub

Private Sub doscan()
Dim c As Long
Dim i As Integer
Dim bx As Integer
Dim by As Integer
Dim x As Integer
Dim y As Integer
Dim r As Integer
Dim sr As Integer
Dim t As Single
Dim dx As Single
Dim dy As Single
Dim te As Single

bx = bot.x * xscale
by = bot.y * yscale + yoffset

i = 0
For t = bot.t - 175 To bot.t + 175 Step 10
    ' introduce random theta error (scan head not pointed correctly) of +/- 4 degrees
    te = t + (Rnd + Rnd + Rnd + Rnd - 2)
    dx = Sin(te / 57.3)
    dy = Cos(te / 57.3)
    scan(i) = 9999
    For r = 5 To 150
        x = bx + r * dx * xscale
        y = by - r * dy * yscale
        c = Form1.Point(x, y)
        Text1 = c
        If c = 0 Then
            ' Add random +/- 2 cm error and quantize to 5 bits
            sr = r / 4 + (Rnd + Rnd + Rnd + Rnd - 2)
            sr = sr * 4
            scan(i) = sr
            'Form1.Show
            x = bx + (r - 4) * dx
            y = by + (r - 4) * dy
            'Form1.Line (bot.x * xscale, bot.y * yscale + yoffset)-(x, y), RGB(0, 255, 0)
            'mark sr * Sin(t / 57.3), sr * Cos(t / 57.3), RGB(0, 0, 255)
            Exit For
        End If
    Next
    i = i + 1
Next

End Sub

Private Sub mark(x As Integer, y As Integer, c As Long)

    Form1.Line (1500 * xscale - 50 + x, 500 * yscale + yoffset + y)-Step(100, 0), c
    Form1.Line (1500 * xscale + x, 500 * yscale + yoffset - 50 + y)-Step(0, 100), c
    
End Sub

Private Sub showbot()
Form1.Circle (bot.x * xscale, bot.y * yscale + yoffset), 50, RGB(0, 255, 0)

End Sub
Private Sub showscan()

' Display scan results in right side of form (robot's eye view of the world)

Dim i As Integer
Dim t As Integer

Form1.Circle (1500 * xscale, 500 * yscale + yoffset), 50, RGB(255, 0, 0)

i = 0
For t = -175 To 175 Step 10
    If scan(i) <> 9999 Then
        mark Sin((t + 180) / 57.3) * scan(i) * xscale, Cos((t + 180) / 57.3) * scan(i) * -yscale, RGB(0, 255, 0)
    Else
        mark Sin((t + 180) / 57.3) * 150 * xscale, Cos((t + 180) / 57.3) * -150 * yscale, RGB(255, 0, 0)
    End If
i = i + 1
Next

End Sub

Private Sub ssbtn_Click()
runme
End Sub

Private Sub Timer1_Timer()

    runme
End Sub
