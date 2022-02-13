VERSION 4.00
Begin VB.Form MainForm 
   Caption         =   "Form1"
   ClientHeight    =   2364
   ClientLeft      =   6432
   ClientTop       =   2052
   ClientWidth     =   5772
   Height          =   2688
   Left            =   6384
   LinkTopic       =   "Form1"
   ScaleHeight     =   2364
   ScaleWidth      =   5772
   Top             =   1776
   Width           =   5868
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   2400
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1680
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   372
      Left            =   2400
      TabIndex        =   0
      Top             =   960
      Width           =   972
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private myrobot As New Robot

Private Sub Command1_Click()

myrobot.accel (-100)
text1.Text = myrobot.Speed()
Call myrobot.place(500, 500)


End Sub


Private Sub Timer1_Timer()
myrobot.accel (20)
End Sub


