VERSION 4.00
Begin VB.Form MasterMind 
   BackColor       =   &H00808080&
   Caption         =   "Master Mind"
   ClientHeight    =   7512
   ClientLeft      =   1080
   ClientTop       =   1680
   ClientWidth     =   8268
   Height          =   7932
   Left            =   1032
   LinkTopic       =   "Form1"
   ScaleHeight     =   7512
   ScaleWidth      =   8268
   Top             =   1308
   Width           =   8364
   Begin VB.Frame Frame2 
      Height          =   492
      Left            =   5760
      TabIndex        =   12
      Top             =   6600
      Width           =   1212
      Begin VB.OptionButton WhiteButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   19
         Top             =   240
         Width           =   252
      End
      Begin VB.OptionButton WhiteButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   4
         Left            =   960
         TabIndex        =   16
         Top             =   240
         Width           =   252
      End
      Begin VB.OptionButton WhiteButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   3
         Left            =   720
         TabIndex        =   15
         Top             =   240
         Width           =   252
      End
      Begin VB.OptionButton WhiteButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   2
         Left            =   480
         TabIndex        =   14
         Top             =   240
         Width           =   252
      End
      Begin VB.OptionButton WhiteButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   252
      End
      Begin VB.Label Label5 
         BackColor       =   &H00808080&
         Caption         =   " 0    1     2    3    4"
         Height          =   252
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1212
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   5760
      TabIndex        =   6
      Top             =   5760
      Width           =   1212
      Begin VB.OptionButton BlackButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   4
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   252
      End
      Begin VB.OptionButton BlackButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   3
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   252
      End
      Begin VB.OptionButton BlackButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   2
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   252
      End
      Begin VB.OptionButton BlackButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   252
      End
      Begin VB.OptionButton BlackButton 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   252
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         Caption         =   " 0    1     2    3    4"
         Height          =   252
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1212
      End
   End
   Begin VB.CommandButton ReadyButton 
      Caption         =   "Ready"
      Height          =   492
      Left            =   5880
      TabIndex        =   5
      Top             =   960
      Width           =   852
   End
   Begin VB.TextBox TurnBox 
      Height          =   288
      Left            =   600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   6840
      Width           =   492
   End
   Begin VB.CommandButton DoneButton 
      Caption         =   "Done"
      Height          =   372
      Left            =   1320
      TabIndex        =   0
      Top             =   6840
      Width           =   972
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Computer's Pattern"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   10.8
         underline       =   0   'False
         italic          =   -1  'True
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   3012
   End
   Begin VB.Shape HideCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   3
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   252
   End
   Begin VB.Shape HideCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   252
   End
   Begin VB.Shape HideCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   252
   End
   Begin VB.Shape HideCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   252
   End
   Begin VB.Image HideImage 
      Height          =   252
      Index           =   3
      Left            =   5280
      Top             =   1080
      Width           =   252
   End
   Begin VB.Image HideImage 
      Height          =   252
      Index           =   2
      Left            =   4800
      Top             =   1080
      Width           =   252
   End
   Begin VB.Image HideImage 
      Height          =   252
      Index           =   1
      Left            =   4320
      Top             =   1080
      Width           =   252
   End
   Begin VB.Image HideImage 
      Height          =   252
      Index           =   0
      Left            =   3840
      Top             =   1080
      Width           =   252
   End
   Begin VB.Line Line9 
      X1              =   3720
      X2              =   5640
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line8 
      X1              =   3720
      X2              =   3720
      Y1              =   1440
      Y2              =   960
   End
   Begin VB.Line Line7 
      X1              =   5640
      X2              =   3720
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line6 
      X1              =   4200
      X2              =   4200
      Y1              =   960
      Y2              =   1440
   End
   Begin VB.Line Line5 
      Index           =   5
      X1              =   4680
      X2              =   4680
      Y1              =   960
      Y2              =   1440
   End
   Begin VB.Line Line5 
      Index           =   4
      X1              =   5160
      X2              =   5160
      Y1              =   960
      Y2              =   1440
   End
   Begin VB.Line Line5 
      Index           =   3
      X1              =   5640
      X2              =   5640
      Y1              =   960
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2040
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   1440
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   2040
      X2              =   120
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line4 
      X1              =   600
      X2              =   600
      Y1              =   960
      Y2              =   1440
   End
   Begin VB.Shape CHideCircle 
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   240
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   252
   End
   Begin VB.Shape CHideCircle 
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   720
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   252
   End
   Begin VB.Line Line5 
      Index           =   0
      X1              =   1080
      X2              =   1080
      Y1              =   960
      Y2              =   1440
   End
   Begin VB.Shape CHideCircle 
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   252
   End
   Begin VB.Line Line5 
      Index           =   1
      X1              =   1560
      X2              =   1560
      Y1              =   960
      Y2              =   1440
   End
   Begin VB.Shape CHideCircle 
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   3
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   252
   End
   Begin VB.Line Line5 
      Index           =   2
      X1              =   2040
      X2              =   2040
      Y1              =   960
      Y2              =   1440
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "White"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   10.8
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4920
      TabIndex        =   3
      Top             =   6720
      Width           =   732
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "Black"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   10.8
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4920
      TabIndex        =   2
      Top             =   5880
      Width           =   732
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   35
      Left            =   240
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   35
      Left            =   240
      Top             =   5400
      Width           =   252
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   35
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   34
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   33
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   32
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   132
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   34
      Left            =   720
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   34
      Left            =   720
      Top             =   5400
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   33
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   33
      Left            =   1200
      Top             =   5400
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   32
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   32
      Left            =   1680
      Top             =   5400
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   28
      Left            =   240
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   28
      Left            =   240
      Top             =   4920
      Width           =   252
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   28
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   29
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   30
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   31
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   132
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   29
      Left            =   720
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   29
      Left            =   720
      Top             =   4920
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   30
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   30
      Left            =   1200
      Top             =   4920
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   31
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   31
      Left            =   1680
      Top             =   4920
      Width           =   252
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   35
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   34
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   33
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   32
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   132
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   35
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   34
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   33
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   32
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   5400
      Width           =   252
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   31
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   30
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   29
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   28
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   132
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   31
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   30
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   29
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   28
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   4920
      Width           =   252
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   27
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   26
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   25
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   24
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   132
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   27
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   26
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   25
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   24
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   252
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   23
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   22
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   21
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   20
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   132
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   23
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   22
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   21
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   20
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   252
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   19
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   18
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   17
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   16
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   132
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   19
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   18
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   17
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   16
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   252
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   15
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   14
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   13
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   12
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   132
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   15
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   14
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   13
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   12
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   252
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   11
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   10
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   9
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   8
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   132
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   11
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   10
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   9
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   8
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   252
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   7
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   6
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   5
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   4
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   132
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   7
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   6
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   5
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   4
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   252
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   3
      Left            =   6600
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   2
      Left            =   6360
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   1
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   132
   End
   Begin VB.Shape CscoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   0
      Left            =   5880
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   132
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   3
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   252
   End
   Begin VB.Shape CguessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   252
   End
   Begin VB.Image ColorChooser 
      Height          =   372
      Index           =   5
      Left            =   3000
      Top             =   6000
      Width           =   372
   End
   Begin VB.Shape ChooseCircle 
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   372
      Index           =   5
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   6000
      Width           =   372
   End
   Begin VB.Image ColorChooser 
      Height          =   372
      Index           =   4
      Left            =   2520
      Top             =   6000
      Width           =   372
   End
   Begin VB.Shape ChooseCircle 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   372
      Index           =   4
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   6000
      Width           =   372
   End
   Begin VB.Image ColorChooser 
      Height          =   372
      Index           =   3
      Left            =   2040
      Top             =   6000
      Width           =   372
   End
   Begin VB.Shape ChooseCircle 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   372
      Index           =   3
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   6000
      Width           =   372
   End
   Begin VB.Image ColorChooser 
      Height          =   372
      Index           =   2
      Left            =   1560
      Top             =   6000
      Width           =   372
   End
   Begin VB.Shape ChooseCircle 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   372
      Index           =   2
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   6000
      Width           =   372
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   27
      Left            =   1680
      Top             =   4440
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   27
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   26
      Left            =   1200
      Top             =   4440
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   26
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   25
      Left            =   720
      Top             =   4440
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   25
      Left            =   720
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   252
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   27
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   26
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   25
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   24
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   132
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   24
      Left            =   240
      Top             =   4440
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   24
      Left            =   240
      Shape           =   3  'Circle
      Top             =   4440
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   23
      Left            =   1680
      Top             =   3960
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   23
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   22
      Left            =   1200
      Top             =   3960
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   22
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   21
      Left            =   720
      Top             =   3960
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   21
      Left            =   720
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   252
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   23
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   22
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   21
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   20
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   132
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   20
      Left            =   240
      Top             =   3960
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   20
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   19
      Left            =   1680
      Top             =   3480
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   19
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   18
      Left            =   1200
      Top             =   3480
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   18
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   17
      Left            =   720
      Top             =   3480
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   17
      Left            =   720
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   252
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   19
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   18
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   17
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   16
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   132
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   16
      Left            =   240
      Top             =   3480
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   16
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   15
      Left            =   1680
      Top             =   3000
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   15
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   14
      Left            =   1200
      Top             =   3000
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   14
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   13
      Left            =   720
      Top             =   3000
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   13
      Left            =   720
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   252
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   15
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   14
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   13
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   12
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   132
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   12
      Left            =   240
      Top             =   3000
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   12
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   11
      Left            =   1680
      Top             =   2520
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   11
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   10
      Left            =   1200
      Top             =   2520
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   10
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   9
      Left            =   720
      Top             =   2520
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   9
      Left            =   720
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   252
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   11
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   10
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   9
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   8
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   132
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   8
      Left            =   240
      Top             =   2520
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   8
      Left            =   240
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   7
      Left            =   1680
      Top             =   2040
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   7
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   6
      Left            =   1200
      Top             =   2040
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   6
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   5
      Left            =   720
      Top             =   2040
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   5
      Left            =   720
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   252
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   7
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   6
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   5
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   4
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   132
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   4
      Left            =   240
      Top             =   2040
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   4
      Left            =   240
      Shape           =   3  'Circle
      Top             =   2040
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   3
      Left            =   1680
      Top             =   1560
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   3
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   2
      Left            =   1200
      Top             =   1560
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   2
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   252
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   1
      Left            =   720
      Top             =   1560
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   1
      Left            =   720
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   252
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   3
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   2
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   1
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   132
   End
   Begin VB.Shape ScoreCircle 
      BorderColor     =   &H00000000&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   132
      Index           =   0
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   132
   End
   Begin VB.Image GuessImage 
      Height          =   252
      Index           =   0
      Left            =   240
      Top             =   1560
      Width           =   252
   End
   Begin VB.Shape GuessCircle 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   252
      Index           =   0
      Left            =   240
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   252
   End
   Begin VB.Shape ShowRect 
      FillStyle       =   0  'Solid
      Height          =   132
      Left            =   600
      Top             =   6480
      Width           =   2772
   End
   Begin VB.Shape ChooseCircle 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   372
      Index           =   1
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   6000
      Width           =   372
   End
   Begin VB.Image ColorChooser 
      Height          =   372
      Index           =   1
      Left            =   1080
      Top             =   6000
      Width           =   372
   End
   Begin VB.Shape ChooseCircle 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   372
      Index           =   0
      Left            =   600
      Shape           =   3  'Circle
      Top             =   6000
      Width           =   372
   End
   Begin VB.Image ColorChooser 
      Height          =   372
      Index           =   0
      Left            =   600
      Top             =   6000
      Width           =   372
   End
End
Attribute VB_Name = "MasterMind"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Const numcolors = 6
Const numpegs = 4

Dim patterns(numcolors, numcolors, numcolors, numcolors) As String * 1
Dim hidden(numpegs) As Integer
Dim Chidden(numpegs) As Integer
Dim guess(numpegs) As Integer
Dim cguess(numpegs) As Integer
Dim Turn As Integer

Dim ChosenColor As Integer
Dim WhiteCount As Integer
Dim Blackcount As Integer

Const BLACK = 0
Const WHITE = &HFFFFFF
Sub ClearPattern()

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer

For i = 0 To numcolors - 1
    For j = 0 To numcolors - 1
        For k = 0 To numcolors - 1
            For l = 0 To numcolors - 1
                patterns(i, j, k, l) = "P"
            Next l
        Next k
    Next j
Next i

End Sub

Sub DisplayGuess()

    Dim i As Integer
    Dim Index As Integer
    
    For i = 0 To numpegs - 1
        Index = GetIndex(Turn, i)
        CguessCircle(Index).FillColor = ChooseCircle(cguess(i)).FillColor
    Next i
    
End Sub

Sub eliminate(g() As Integer, b As Integer, w As Integer)

' g() contains a guess resulting in b and w
' eliminate all combinations that don't match

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim h(numpegs) As Integer
    
    For i = 0 To numcolors - 1
        For j = 0 To numcolors - 1
            For k = 0 To numcolors - 1
                For l = 0 To numcolors - 1
                    If (patterns(i, j, k, l) = "P") Then
                        h(0) = i
                        h(1) = j
                        h(2) = k
                        h(3) = l
                        Call score(h(), g())
                        If (Blackcount <> b) Or (WhiteCount <> w) Then
                            patterns(i, j, k, l) = "N"
                        End If
                    End If
                Next l
            Next k
        Next j
    Next i
    
End Sub
Sub GetGuess()
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    
    Dim ix As Integer
    Dim jx As Integer
    Dim kx As Integer
    Dim lx As Integer
   
    ix = Int(Rnd(1) * numcolors)
    jx = Int(Rnd(1) * numcolors)
    kx = Int(Rnd(1) * numcolors)
    lx = Int(Rnd(1) * numcolors)
    
    For i = ix To numcolors - 1
        For j = jx To numcolors - 1
            For k = kx To numcolors - 1
                For l = lx To numcolors - 1
                    If (patterns(i, j, k, l) = "P") Then
                        cguess(0) = i
                        cguess(1) = j
                        cguess(2) = k
                        cguess(3) = l
                        Exit Sub
                    End If
                Next l
            Next k
        Next j
    Next i

    For i = 0 To ix
        For j = 0 To jx
            For k = 0 To kx
                For l = 0 To lx
                    If (patterns(i, j, k, l) = "P") Then
                        cguess(0) = i
                        cguess(1) = j
                        cguess(2) = k
                        cguess(3) = l
                        Exit Sub
                    End If
                Next l
            Next k
        Next j
    Next i

Debug.Print "Cheat Cheat never beat"

End Sub

Sub score(h() As Integer, g() As Integer)

    Dim i As Integer
    Dim j As Integer
    Dim Index As Integer
    
    Dim hiddenx(numpegs) As Integer
    Dim guessx(numpegs) As Integer
    
    ' make a copy
    For i = 0 To numpegs - 1
        guessx(i) = g(i)
        hiddenx(i) = h(i)
    Next i
    
    ' count blacks
    Blackcount = 0
    For i = 0 To numpegs - 1
        If hiddenx(i) = guessx(i) Then
            Blackcount = Blackcount + 1
            hiddenx(i) = -1
            guessx(i) = -2
        End If
    Next i
    
    ' count whites
    WhiteCount = 0
    For i = 0 To numpegs - 1
        For j = 0 To numpegs - 1
           If hiddenx(i) = guessx(j) Then
                WhiteCount = WhiteCount + 1
                hiddenx(i) = -1
                guessx(j) = -2
           End If
        Next j
    Next i
    
    
End Sub

Function GetIndex(t As Integer, i As Integer)

GetIndex = t * numpegs + i

End Function



Sub MakeHidden()

    Dim i As Integer
    
    For i = 0 To numpegs - 1
        Chidden(i) = Int(Rnd(1) * numcolors)
    Next i
    
End Sub

Sub ShowHidden()

    Dim i As Integer

    For i = 0 To numpegs - 1
        CHideCircle(i).FillColor = ChooseCircle(Chidden(i)).FillColor
    Next i
    
End Sub


Private Sub ColorChooser_Click(Index As Integer)

ShowRect.FillColor = ChooseCircle(Index).FillColor
ChosenColor = Index

End Sub


Private Sub DoneButton_Click()

    Dim i As Integer
    Dim Index As Integer
    Dim b As Integer
    Dim w As Integer
    
    Turn = Val(TurnBox.Text)
    
    Call score(Chidden(), guess())
    For i = 0 To (Blackcount - 1)
        Index = GetIndex(Turn, i)
        MasterMind.ScoreCircle(Index).FillColor = BLACK
    Next i
    
    For i = Blackcount To (Blackcount + WhiteCount - 1)
        Index = GetIndex(Turn, i)
        MasterMind.ScoreCircle(Index).FillColor = WHITE
    Next i
   
    For i = 0 To numpegs
        If BlackButton(i).Value = True Then
            b = i
            Exit For
        End If
    Next i
    
    For i = 0 To numpegs
        If WhiteButton(i).Value = True Then
            w = i
            Exit For
        End If
    Next i
    Debug.Print cguess(0), cguess(1), cguess(2), cguess(3)
    Debug.Print b, w
    
    Call eliminate(cguess(), b, w)
    
    Turn = Turn + 1
    
    TurnBox.Text = Str$(Turn)
    GetGuess
    DisplayGuess
    If Blackcount = numpegs Then
        ShowHidden
    End If
    
End Sub




Private Sub GuessImage_Click(Index As Integer)
    
    GuessCircle(Index).FillColor = ShowRect.FillColor
    guess(Index - Turn * numpegs) = ChosenColor
 
End Sub





Private Sub Form_Load()

    Randomize
    ClearPattern
    MakeHidden
    TurnBox.Text = "0"
    
End Sub









Private Sub HideImage_Click(Index As Integer)

    HideCircle(Index).FillColor = ChooseCircle(ChosenColor).FillColor
    hidden(Index) = ChosenColor
    Debug.Print "Hidden "; Index; " is color "; ChosenColor
End Sub


Private Sub ReadyButton_Click()

    GetGuess
    DisplayGuess
    
End Sub




