VERSION 2.00
Begin Form ComForm 
   Caption         =   "Communications"
   ClientHeight    =   1716
   ClientLeft      =   888
   ClientTop       =   1560
   ClientWidth     =   3840
   Height          =   2136
   Left            =   840
   LinkTopic       =   "Form1"
   ScaleHeight     =   1716
   ScaleWidth      =   3840
   Top             =   1188
   Width           =   3936
   Begin CommandButton RecvButton 
      Caption         =   "Recieve"
      Height          =   372
      Left            =   1800
      TabIndex        =   5
      Top             =   1200
      Width           =   972
   End
   Begin CommandButton SendBtn 
      Caption         =   "Send"
      Height          =   372
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   972
   End
   Begin TextBox Text4 
      Height          =   372
      Left            =   0
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   1200
      Width           =   1572
   End
   Begin TextBox Text3 
      Height          =   372
      Left            =   0
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   720
      Width           =   1572
   End
   Begin Timer Timer1 
      Interval        =   250
      Left            =   3480
      Top             =   360
   End
   Begin TextBox Text2 
      Height          =   288
      Left            =   0
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   360
      Width           =   3492
   End
   Begin TextBox Text1 
      Height          =   288
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   3492
   End
   Begin MSComm Comm1 
      CommPort        =   2
      Handshaking     =   1  'XON/XOFF
      Interval        =   1000
      Left            =   3480
      Top             =   0
   End
End
Option Explicit

Sub Form_Load ()

    OpenCom

End Sub

Sub RecvButton_Click ()

'     text4.Text = GetCom()

    PsuedoGet

End Sub

Sub SendBtn_Click ()

    Dim buffer As String

    buffer = text3.Text
    SendCom buffer

End Sub

Sub Timer1_Timer ()
    LoadCommBuff
End Sub

