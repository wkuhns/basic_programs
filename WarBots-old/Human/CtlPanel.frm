VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4490
   ClientLeft      =   -20
   ClientTop       =   5820
   ClientWidth     =   6260
   LinkTopic       =   "Form1"
   ScaleHeight     =   4490
   ScaleWidth      =   6260
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   372
      Left            =   4200
      TabIndex        =   20
      Top             =   360
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Height          =   372
      Left            =   4080
      TabIndex        =   15
      Top             =   3120
      Width           =   132
   End
   Begin VB.CommandButton Command2 
      Height          =   372
      Left            =   5640
      TabIndex        =   14
      Top             =   3120
      Width           =   132
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   5760
      Top             =   2520
   End
   Begin VB.PictureBox Picture1 
      Height          =   3612
      Left            =   120
      ScaleHeight     =   2018.388
      ScaleMode       =   0  'User
      ScaleTop        =   999
      ScaleWidth      =   1953.458
      TabIndex        =   12
      Top             =   120
      Width           =   3732
   End
   Begin VB.TextBox Text2 
      Height          =   288
      Left            =   4320
      TabIndex        =   10
      Top             =   1080
      Width           =   492
   End
   Begin VB.OptionButton Option1 
      Height          =   252
      Index           =   7
      Left            =   3120
      TabIndex        =   8
      Top             =   3960
      Width           =   252
   End
   Begin VB.OptionButton Option1 
      Height          =   252
      Index           =   6
      Left            =   2760
      TabIndex        =   7
      Top             =   3960
      Width           =   252
   End
   Begin VB.OptionButton Option1 
      Height          =   252
      Index           =   5
      Left            =   2400
      TabIndex        =   6
      Top             =   3960
      Width           =   252
   End
   Begin VB.OptionButton Option1 
      Height          =   252
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   3960
      Width           =   252
   End
   Begin VB.OptionButton Option1 
      Height          =   252
      Index           =   3
      Left            =   1680
      TabIndex        =   4
      Top             =   3960
      Width           =   252
   End
   Begin VB.OptionButton Option1 
      Height          =   252
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   3960
      Width           =   252
   End
   Begin VB.OptionButton Option1 
      Height          =   252
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   3960
      Width           =   252
   End
   Begin VB.OptionButton Option1 
      Height          =   252
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   3960
      Width           =   252
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   4080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3840
      Width           =   1932
   End
   Begin ComctlLib.Slider SetSpeedSlider 
      Height          =   504
      Left            =   4200
      TabIndex        =   17
      Top             =   3120
      Width           =   1452
      _ExtentX        =   2558
      _ExtentY        =   741
      _Version        =   327680
      MouseIcon       =   "CtlPanel.frx":0000
      SmallChange     =   5
      Max             =   100
      TickStyle       =   1
      TickFrequency   =   10
   End
   Begin ComctlLib.Slider ActSpeedSlider 
      Height          =   492
      Left            =   4200
      TabIndex        =   18
      Top             =   2760
      Width           =   1452
      _ExtentX        =   2558
      _ExtentY        =   741
      _Version        =   327680
      MouseIcon       =   "CtlPanel.frx":001C
      SmallChange     =   5
      Max             =   100
      TickFrequency   =   10
   End
   Begin VB.Shape Shape2 
      Height          =   252
      Left            =   4920
      Top             =   1080
      Width           =   612
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   132
      Left            =   120
      Top             =   0
      Width           =   3732
   End
   Begin VB.Label Label1 
      Caption         =   " 1        2        3       4        5       6        8      10"
      Height          =   252
      Left            =   600
      TabIndex        =   9
      Top             =   4200
      Width           =   2772
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' MyForm is the startup object for this application.
' This form (MyForm) is not visible at run-time. It is only required
' because of object referencing limitations in Visual Basic
' Client/Server applications.

' There are a small number of required subroutines. Some should
' not be altered:
'    Form_Load - establishes the linkage to the robot server
'    Die - Used to clean up and exit

' There are three subroutines that are automatically run by
' the robot, which may contain user code:
'    UserInit - executed once at startup.
'    Ping - executed by the server when this robot is scanned
'       by another robot
'    UserFrame - run continuously as long as the robot is
'       alive.
' Any other user created subroutines must be called from one
' of these.

' There is one required global object: MyRobot.
' This is the robot object which provides the interface to
' the robot server. In VB5, you can use the Object Browser
' (F2 Key) to view the methods available.

Dim MyRobot As RobotDLL.Robot

' User defined globals:
' These are 'global' to this form. Use these or add your own.
' They are not required except as used by your application

Dim speed As Integer
Dim scandir As Single
Dim dir As Integer
Dim scanres As Integer
Dim theta As Integer
Dim cleartime As Long
Dim ticks As Long
'
' Our life is over. Release robot's soul and die. Do not
' change this subroutine
'
Sub Die()

Set MyRobot = Nothing

End

End Sub

'
' This subroutine MUST exist with EXACTLY this name and
' this argument list. The contents are up to the user.
' This subroutine is invoked by the server when this robot
' is scanned by another robot.
'
Public Sub Ping(m As String)

    Shape1.FillColor = RGB(255, 0, 0)
    Text2.Text = m
    cleartime = ticks + 4
    
End Sub





Sub UserInit()

dir = 0
speed = 100
MyRobot.SetName ("Human")
Timer1.Enabled = True
Picture1.Scale (0, 999)-(999, 0)
Option1_Click 7

End Sub

Private Sub Command1_Click()

SetSpeedSlider.Value = 35
MyRobot.Drive theta, 35

End Sub

Private Sub Command2_Click()

SetSpeedSlider.Value = 100
MyRobot.Drive theta, 100

End Sub

Private Sub Command3_Click()

Picture1.Cls

End Sub

'
' Don't change this at all. This code creates the linkage
' to the robot server process.
'
Private Sub Form_Load()

' Create robot object

Set MyRobot = CreateObject("RobotDLL.Robot")

' Register 'Ping' procedure with server.

Call MyRobot.RegisterAlert(Form1)

' Do user's initialization.

UserInit

' Don't change this - User specific stuff is in DoFrame.

End Sub

Private Sub Form_Unload(Cancel As Integer)

Die

End Sub
Private Sub Gauge1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim cx As Integer
Dim cy As Integer
Dim dx As Integer
Dim dy As Integer

cx = Gauge1.Width / 2
cy = Gauge1.Height / 2

dx = X - cx
dy = cy - Y
If dx = 0 Then dx = 1
theta = Atn(dy / dx) * 57.3
If dx < 0 Then theta = theta + 180
If theta < 0 Then theta = theta + 360

Text1 = Str(theta)

Gauge1.Value = ((360 - theta) + 180) Mod 360
ActDirGauge.Value = Gauge1.Value
MyRobot.Drive theta, SetSpeedSlider.Value

End Sub

Private Sub Option1_Click(Index As Integer)

Select Case Index
    Case 0: scanres = 1
    Case 1: scanres = 2
    Case 2: scanres = 3
    Case 3: scanres = 4
    Case 4: scanres = 5
    Case 5: scanres = 6
    Case 6: scanres = 8
    Case 7: scanres = 10
End Select

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim cx As Integer
Dim cy As Integer
Dim dx As Double
Dim dy As Double
Dim mytheta As Single
Dim range As Integer
Dim rsq As Double

dx = X - MyRobot.X
dy = Y - MyRobot.Y
If dx = 0 Then dx = 1
mytheta = Atn(dy / dx) * 57.3
If dx < 0 Then mytheta = mytheta + 180
If mytheta < 0 Then mytheta = mytheta + 360

If Button = 1 Then
    range = MyRobot.scan(mytheta, scanres)
    
    If range > 0 Then
        On Error Resume Next
        Picture1.Circle (MyRobot.X, MyRobot.Y), range, 0, (mytheta - scanres) / 57.3, (mytheta + scanres) / 57.3
        Text1.Text = "Bogey at " + Str(range)
    Else
        Text1.Text = ""
    End If
End If

If Button = 2 Then
    rsq = (dx * dx + dy * dy)
    range = Sqr(rsq)
    If MyRobot.cannon(Int(mytheta), range) = -1 Then
        Text1.Text = "Bang"
    Else
        Text1.Text = "Click"
    End If
End If
    
End Sub


Private Sub Timer1_Timer()

    Static X As Integer
    Static Y As Integer
    
    ticks = ticks + 1
    If ticks > cleartime Then
        Shape1.FillColor = &HC0C0C0
        Text2.Text = ""
    End If
    
    ' Check to see if we're dead. You can't cheat death
    ' by changing this - all that will happen is that
    ' you'll have dead processes cluttering up your
    ' system.
    If MyRobot.status = "K" Then
        Die
        Exit Sub
    End If
    ' erase old box
    Picture1.Line (X - 10, Y - 10)-Step(20, 20), Picture1.BackColor, BF
    
    ' ShowStatus MUST be called periodically.
    MyRobot.ShowStatus
    X = MyRobot.X
    Y = MyRobot.Y
'    Text1.Text = Str(X) + "," + Str(Y)
    Picture1.FillStyle = 0
    Picture1.Line (X - 10, Y - 10)-Step(20, 20), RGB(255, 0, 0), BF
    Gauge2.Value = MyRobot.heat
    Gauge3.Value = MyRobot.health
    
    ActSpeedSlider.Value = MyRobot.speed
'    If ActSpeedSlider.Value <> SetSpeedSlider.Value Then
        MyRobot.Drive theta, SetSpeedSlider.Value
'    End If
    
End Sub


