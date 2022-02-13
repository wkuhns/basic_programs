VERSION 2.00
Begin Form ColorForm 
   Caption         =   "Colors"
   ClientHeight    =   4200
   ClientLeft      =   876
   ClientTop       =   1524
   ClientWidth     =   3852
   Height          =   4620
   Left            =   828
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   3852
   Top             =   1152
   Width           =   3948
   Begin TextBox Text1 
      Height          =   372
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   972
   End
   Begin PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3852
      Left            =   0
      Picture         =   COLORFOR.FRX:0000
      ScaleHeight     =   3828
      ScaleWidth      =   3828
      TabIndex        =   0
      Top             =   360
      Width           =   3852
   End
End
Option Explicit

Sub Picture1_MouseDown (Button As Integer, Shift As Integer, x As Single, y As Single)

    text1.Text = Hex$(picture1.Point(x, y))

End Sub

