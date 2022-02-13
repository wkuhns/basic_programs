VERSION 5.00
Begin VB.Form StatForm 
   Caption         =   "Form1"
   ClientHeight    =   1428
   ClientLeft      =   7632
   ClientTop       =   276
   ClientWidth     =   4632
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1428
   ScaleWidth      =   4632
   Begin VB.TextBox PostBox 
      DragMode        =   1  'Automatic
      Height          =   1212
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "status.frx":0000
      Top             =   0
      Width           =   3132
   End
   Begin VB.TextBox HeatBox 
      Height          =   288
      Left            =   720
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   720
      Width           =   612
   End
   Begin VB.TextBox HealthBox 
      Height          =   288
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   612
   End
   Begin VB.TextBox DirBox 
      Height          =   288
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   612
   End
   Begin VB.TextBox SpeedBox 
      Height          =   288
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   612
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   132
      Left            =   1440
      Top             =   1200
      Width           =   3132
   End
   Begin VB.Label Label5 
      Caption         =   "Heat"
      Height          =   252
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   732
   End
   Begin VB.Label Label4 
      Caption         =   "Health"
      Height          =   252
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "Direction"
      Height          =   252
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "Speed"
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   972
   End
End
Attribute VB_Name = "StatForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit











Private Sub Form_Resize()

PostBox.Text = "Resized " + Str(StatForm.Height) + " " + Str(StatForm.ScaleHeight)
PostBox.Width = StatForm.Width - 1600
PostBox.Height = StatForm.Height - 540
Shape1.Top = PostBox.Height

End Sub




Private Sub PostBox_Change()

'PostBox.Text = "Changed " + Str(StatForm.ScaleHeight) + " " + Str(StatForm.ScaleHeight)
'PostBox.Width = StatForm.Width - 1600
'PostBox.Height = StatForm.Height - 540
'Shape1.Top = PostBox.Height

End Sub


