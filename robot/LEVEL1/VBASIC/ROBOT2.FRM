VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   4044
   ClientLeft      =   4812
   ClientTop       =   5292
   ClientWidth     =   5688
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   7.8
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Height          =   4464
   Left            =   4764
   LinkTopic       =   "Form1"
   Picture         =   "ROBOT2.frx":0000
   ScaleHeight     =   4044
   ScaleWidth      =   5688
   Top             =   4920
   Width           =   5784
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4920
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox J3AngleBox 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   2760
      TabIndex        =   10
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox J3SensorBox 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   2760
      TabIndex        =   9
      Top             =   480
      Width           =   972
   End
   Begin VB.TextBox J2AngleBox 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1440
      TabIndex        =   8
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox J2SensorBox 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1440
      TabIndex        =   7
      Top             =   480
      Width           =   972
   End
   Begin VB.TextBox J1AngleBox 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox J1SensorBox 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   972
   End
   Begin VB.PictureBox Canvas2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00808080&
      Height          =   1812
      Left            =   2760
      Picture         =   "ROBOT2.frx":0446
      ScaleHeight     =   150
      ScaleMode       =   0  'User
      ScaleWidth      =   200
      TabIndex        =   4
      Top             =   1200
      Width           =   2532
   End
   Begin VB.HScrollBar Joint3 
      Height          =   252
      Left            =   2760
      Max             =   270
      TabIndex        =   3
      Top             =   3120
      Value           =   135
      Width           =   2532
   End
   Begin VB.HScrollBar Joint2 
      Height          =   252
      Left            =   120
      Max             =   90
      TabIndex        =   2
      Top             =   3480
      Value           =   45
      Width           =   2532
   End
   Begin VB.HScrollBar Joint1 
      Height          =   252
      Left            =   120
      Max             =   180
      TabIndex        =   1
      Top             =   3120
      Value           =   90
      Width           =   2532
   End
   Begin VB.PictureBox Canvas 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FF0000&
      Height          =   1812
      Left            =   120
      Picture         =   "ROBOT2.frx":7DB8
      ScaleHeight     =   150
      ScaleMode       =   0  'User
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   1200
      Width           =   2532
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub DrawRobot()
    Canvas.Cls
    Canvas2.Cls
    
    Canvas.DrawWidth = 11
    Canvas.Line (175 - j(2).x, 175 - j(2).y)-(175 - j(3).x, 175 - j(3).y), RGB(0, 0, 255)
    
    Canvas2.DrawWidth = 11
    Canvas2.Line (j(2).x, j(2).z)-(j(3).x, j(3).z), RGB(0, 0, 255)

    Canvas.DrawWidth = 3
    Canvas.Line (175 - j(3).x, 175 - j(3).y)-(175 - j(4).x, 175 - j(4).y), RGB(255, 0, 0)
    
    Canvas2.DrawWidth = 3
    Canvas2.Line (j(3).x, j(3).z)-(j(4).x, j(4).z), RGB(255, 0, 0)
    
    J1AngleBox = j(1).theta
    J1SensorBox = j(1).s
    J2AngleBox = j(2).theta
    J2SensorBox = j(2).s
    J3AngleBox = j(3).theta
    J3SensorBox = j(3).s

End Sub

Private Sub Form_Load()
    
    j(1).x = 75
    j(1).y = 75
    j(1).z = 60
    j(1).r = 10

    j(2).z = 70
    j(2).r = 50

    j(3).r = 50

    For i = 1 To 4
        ReadCalFile (i)
    Next i

End Sub

Private Sub Joint1_Change()
    j(1).theta = Joint1.Value / 57.29578
    J1AngleBox = Joint1.Value
    resolve
    DrawRobot
End Sub

Private Sub Joint2_Change()
    j(2).theta = 0.7854 - (Joint2.Value / 57.29578)
    J2AngleBox = Joint2.Value
    resolve
    DrawRobot

End Sub

Private Sub Joint3_Change()
    j(3).theta = 2.25 - (Joint3.Value / 57.29578)
    J3AngleBox = Joint3.Value
    resolve
    DrawRobot

End Sub

Private Sub Timer1_Timer()
    ReadJoysticks
    For i = 1 To 4
        GetDegrees (i)
    Next i
    resolve
    DrawRobot
End Sub

