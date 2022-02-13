VERSION 5.00
Begin VB.Form PlayerForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Players"
   ClientHeight    =   2860
   ClientLeft      =   5700
   ClientTop       =   2760
   ClientWidth     =   4370
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
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2860
   ScaleWidth      =   4370
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   372
      Left            =   3240
      TabIndex        =   16
      Top             =   2280
      Width           =   972
   End
   Begin VB.CheckBox RemoteBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   3720
      TabIndex        =   14
      Top             =   1800
      Width           =   252
   End
   Begin VB.CheckBox RemoteBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   3720
      TabIndex        =   13
      Top             =   1440
      Width           =   252
   End
   Begin VB.CheckBox RemoteBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   3720
      TabIndex        =   12
      Top             =   1080
      Width           =   252
   End
   Begin VB.CheckBox RemoteBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   3720
      TabIndex        =   11
      Top             =   720
      Width           =   252
   End
   Begin VB.CheckBox RemoteBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   3720
      TabIndex        =   10
      Top             =   360
      Width           =   252
   End
   Begin VB.TextBox PlayerName 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   4
      Left            =   1560
      TabIndex        =   9
      Top             =   1800
      Width           =   1572
   End
   Begin VB.TextBox PlayerName 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   3
      Left            =   1560
      TabIndex        =   8
      Top             =   1440
      Width           =   1572
   End
   Begin VB.TextBox PlayerName 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   2
      Left            =   1560
      TabIndex        =   7
      Top             =   1080
      Width           =   1572
   End
   Begin VB.TextBox PlayerName 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   1
      Left            =   1560
      TabIndex        =   6
      Top             =   720
      Width           =   1572
   End
   Begin VB.TextBox PlayerName 
      Appearance      =   0  'Flat
      Height          =   288
      Index           =   0
      Left            =   1560
      TabIndex        =   5
      Top             =   360
      Width           =   1572
   End
   Begin VB.CheckBox PlayerBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Player 5"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   972
   End
   Begin VB.CheckBox PlayerBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Player 4"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   972
   End
   Begin VB.CheckBox PlayerBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Player 3"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   972
   End
   Begin VB.CheckBox PlayerBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Player 2"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   972
   End
   Begin VB.CheckBox PlayerBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Player 1"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   972
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Remote"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3480
      TabIndex        =   15
      Top             =   120
      Width           =   732
   End
End
Attribute VB_Name = "PlayerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    PlayerForm.Hide                 ' unload?

End Sub

Private Sub Form_Load()
    
    Dim i As Integer

    For i = 0 To 4
        If actives(i) Then
            PlayerBox(i).Value = 1
        Else
            PlayerBox(i).Value = 0
        End If
        PlayerName(i).Text = cars(i).name
        If cars(i).local = 0 Then
            RemoteBox(i).Value = True
        Else
            RemoteBox(i).Value = False
        End If

    Next i

End Sub

Private Sub PlayerBox_Click(index As Integer)
    
    If PlayerBox(index).Value = 0 Then
        cars(index).active = False
        actives(index) = False
    Else
        cars(index).active = True
        actives(index) = True
    End If

End Sub

Private Sub PlayerName_Change(index As Integer)
    cars(index).name = PlayerName(index).Text
End Sub

Private Sub RemoteBox_Click(index As Integer)
    
    If RemoteBox(index).Value <> 0 Then
        cars(index).local = False
    Else
        cars(index).local = True
    End If

End Sub

