VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form Server 
   Caption         =   "Server"
   ClientHeight    =   1908
   ClientLeft      =   1548
   ClientTop       =   5532
   ClientWidth     =   5868
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1908
   ScaleWidth      =   5868
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   120
      TabIndex        =   16
      Top             =   1440
      Width           =   5652
   End
   Begin VB.CommandButton DebugBtn 
      Caption         =   "Debug"
      Height          =   372
      Left            =   2640
      TabIndex        =   15
      Top             =   480
      Width           =   972
   End
   Begin VB.CommandButton RefreshBtn1 
      Caption         =   "Redraw"
      Height          =   372
      Left            =   3720
      TabIndex        =   14
      Top             =   480
      Width           =   972
   End
   Begin VB.CommandButton QuitBtn 
      Caption         =   "Quit"
      Height          =   372
      Left            =   4800
      TabIndex        =   13
      Top             =   960
      Width           =   972
   End
   Begin ComctlLib.Slider HealthBar 
      Height          =   132
      Index           =   0
      Left            =   1080
      TabIndex        =   9
      Top             =   120
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   233
      _Version        =   327680
      MouseIcon       =   "Server.frx":0000
      Max             =   100
      TickFrequency   =   10
   End
   Begin VB.CommandButton ResetBtn 
      Caption         =   "Junk These      Robots"
      Height          =   372
      Left            =   4800
      TabIndex        =   8
      Top             =   480
      Width           =   972
   End
   Begin VB.CommandButton RestartBtn 
      Caption         =   "Restart"
      Height          =   372
      Left            =   4800
      TabIndex        =   7
      Top             =   0
      Width           =   972
   End
   Begin VB.TextBox StatBox 
      Height          =   288
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Text            =   "Null"
      Top             =   1080
      Width           =   972
   End
   Begin VB.TextBox StatBox 
      Height          =   288
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Text            =   "Null"
      Top             =   720
      Width           =   972
   End
   Begin VB.TextBox StatBox 
      Height          =   288
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Text            =   "Null"
      Top             =   360
      Width           =   972
   End
   Begin VB.TextBox StatBox 
      Height          =   288
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Text            =   "Null"
      Top             =   0
      Width           =   972
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   2640
      TabIndex        =   2
      Text            =   "Ready"
      Top             =   960
      Width           =   2052
   End
   Begin VB.CommandButton PauseBtn 
      Caption         =   "Pause"
      Height          =   372
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   972
   End
   Begin VB.CommandButton RunBtn 
      Caption         =   "Run"
      Height          =   372
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   972
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2400
      Top             =   240
   End
   Begin ComctlLib.Slider HealthBar 
      Height          =   132
      Index           =   1
      Left            =   1080
      TabIndex        =   10
      Top             =   480
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   233
      _Version        =   327680
      MouseIcon       =   "Server.frx":001C
      Max             =   100
      TickFrequency   =   10
   End
   Begin ComctlLib.Slider HealthBar 
      Height          =   132
      Index           =   2
      Left            =   1080
      TabIndex        =   11
      Top             =   840
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   233
      _Version        =   327680
      MouseIcon       =   "Server.frx":0038
      Max             =   100
      TickFrequency   =   10
   End
   Begin ComctlLib.Slider HealthBar 
      Height          =   132
      Index           =   3
      Left            =   1080
      TabIndex        =   12
      Top             =   1200
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   233
      _Version        =   327680
      MouseIcon       =   "Server.frx":0054
      Max             =   100
      TickFrequency   =   10
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DebugBtn_Click()

If DebugState Then
    DebugState = False
    DebugBtn.Caption = "Debug"
Else
    DebugState = True
    DebugBtn.Caption = "End Debug"
End If

End Sub


Private Sub Form_Load()

status = "P"

Bots(1).color = RGB(255, 0, 0)
Bots(2).color = RGB(0, 255, 0)
Bots(3).color = RGB(0, 0, 255)
Bots(4).color = RGB(255, 0, 255)
Randomize

Arena.Visible = True

End Sub

Public Sub PauseBtn_Click()

Timer1.Enabled = False
Text2 = "Paused"
status = "P"

End Sub

Private Sub QuitBtn_Click()

Dim finis As Long

finis = Timer + 4
Server.Text2 = "Cleaning up...."
ResetBtn_Click

While Timer < finis
    DoEvents
    Sleep 100
Wend

End

End Sub

Private Sub RefreshBtn1_Click()

Arena.Cls

End Sub

Public Sub ResetBtn_Click()

Dim i As Integer

Timer1.Enabled = False

For i = 1 To LastBot
    ' Clean up structures
    KillBot (i)
    StatBox(i - 1).BackColor = Text2.BackColor
    StatBox(i - 1).Text = "Null"
    ' This next step generates an error, since
    ' it kills the client process, which then fails to
    ' complete the OLE handshake.
    On Error Resume Next
    Call Bots(i).proc.die
Next i

Arena.Form_Load

Text2 = "Reset, Paused"
 
LastBot = 0

End Sub

Private Sub RestartBtn_Click()

Dim i As Integer

PauseBtn_Click
Arena.Form_Load

For i = 1 To LastBot
    KillBot (i)
Next i

For i = 1 To LastBot
    PlaceBot (i)
    StatBox(i - 1).Text = "Ready"
Next i

Text2.Text = "Restarted, Paused"

End Sub

Public Sub RunBtn_Click()

Timer1.Enabled = True
Text2 = "Running"
status = "R"

End Sub


Private Sub Timer1_Timer()

MoveBots
Arena.DrawFrame

tick = tick + 1

End Sub


