VERSION 4.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Robot 1"
   ClientHeight    =   1992
   ClientLeft      =   1128
   ClientTop       =   1560
   ClientWidth     =   6420
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
   Height          =   2412
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   1992
   ScaleWidth      =   6420
   Top             =   1188
   Width           =   6516
   Begin VB.TextBox ch4box 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   5280
      TabIndex        =   4
      Top             =   720
      Width           =   972
   End
   Begin VB.TextBox ch3box 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   4080
      TabIndex        =   3
      Top             =   720
      Width           =   972
   End
   Begin VB.TextBox ch2box 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   2880
      TabIndex        =   2
      Top             =   720
      Width           =   972
   End
   Begin VB.TextBox ch1box 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Read"
      Height          =   492
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   972
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Channel 4"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5280
      TabIndex        =   8
      Top             =   480
      Width           =   972
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Channel 3"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   4080
      TabIndex        =   7
      Top             =   480
      Width           =   972
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Channel 2"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2880
      TabIndex        =   6
      Top             =   480
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Channel 1"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim a, b, c, d As Integer
    Call ReadJoysticks(a, b, c, d)
    ch1box.Text = a
    ch2box.Text = b
    ch3box.Text = c
    ch4box.Text = d
End Sub

