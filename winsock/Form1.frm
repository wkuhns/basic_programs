VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3201
   ClientLeft      =   55
   ClientTop       =   341
   ClientWidth     =   4686
   LinkTopic       =   "Form1"
   ScaleHeight     =   3201
   ScaleWidth      =   4686
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton SendBtn 
      Caption         =   "Send"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox RecvBox 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox SendBox 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton ListenBtn 
      Caption         =   "Listen"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   480
      Top             =   600
      _ExtentX        =   547
      _ExtentY        =   547
      _Version        =   393216
      LocalPort       =   713
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ListenBtn_Click()

    Winsock.Listen
    
End Sub

Private Sub SendBtn_Click()

    Winsock.SendData (SendBox.Text)
    
End Sub

Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)

   If Winsock.State <> sckClosed Then Winsock.Close

   ' Pass the value of the requestID parameter to the
   ' Accept method.
   Winsock.Accept requestID

    
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim strdata As String
    Winsock.GetData strdata, vbString
    RecvBox.Text = strdata
    
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    MsgBox "Winsock Error"
    
End Sub
