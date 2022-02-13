VERSION 5.00
Begin VB.Form ScoreBoard 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Scoreboard"
   ClientHeight    =   1300
   ClientLeft      =   6130
   ClientTop       =   1620
   ClientWidth     =   3590
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1300
   ScaleWidth      =   3590
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1120
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3612
   End
End
Attribute VB_Name = "ScoreBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
