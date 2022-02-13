VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form ServerForm 
   Caption         =   "Server"
   ClientHeight    =   1416
   ClientLeft      =   1548
   ClientTop       =   5532
   ClientWidth     =   5868
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1416
   ScaleWidth      =   5868
   Begin VB.CheckBox SoundBox 
      Height          =   253
      Left            =   2299
      TabIndex        =   16
      Top             =   242
      Value           =   1  'Checked
      Width           =   253
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
      Left            =   1078
      TabIndex        =   9
      Top             =   121
      Width           =   1210
      _ExtentX        =   2138
      _ExtentY        =   233
      _Version        =   327682
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
      Left            =   2299
      TabIndex        =   2
      Text            =   "Ready"
      Top             =   960
      Width           =   2398
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
      Interval        =   20
      Left            =   2783
      Top             =   726
   End
   Begin ComctlLib.Slider HealthBar 
      Height          =   132
      Index           =   1
      Left            =   1078
      TabIndex        =   10
      Top             =   484
      Width           =   1210
      _ExtentX        =   2138
      _ExtentY        =   233
      _Version        =   327682
      Max             =   100
      TickFrequency   =   10
   End
   Begin ComctlLib.Slider HealthBar 
      Height          =   132
      Index           =   2
      Left            =   1078
      TabIndex        =   11
      Top             =   836
      Width           =   1210
      _ExtentX        =   2138
      _ExtentY        =   233
      _Version        =   327682
      Max             =   100
      TickFrequency   =   10
   End
   Begin ComctlLib.Slider HealthBar 
      Height          =   132
      Index           =   3
      Left            =   1078
      TabIndex        =   12
      Top             =   1199
      Width           =   1210
      _ExtentX        =   2138
      _ExtentY        =   233
      _Version        =   327682
      Max             =   100
      TickFrequency   =   10
   End
End
Attribute VB_Name = "ServerForm"
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

rsinitialize

End Sub

Public Sub PauseBtn_Click()

pause

End Sub

Private Sub QuitBtn_Click()

CleanUpAndDie

End Sub

Private Sub RefreshBtn1_Click()

arena.Cls

End Sub

Public Sub ResetBtn_Click()

reset

End Sub

Private Sub RestartBtn_Click()

restart

End Sub

Public Sub RunBtn_Click()

run

End Sub


Private Sub Timer1_Timer()

rstick

End Sub


