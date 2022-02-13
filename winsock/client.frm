VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Client 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
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
   Begin VB.CommandButton ConnectBtn 
      Caption         =   "Connect"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   480
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "wiley"
      RemotePort      =   713
      LocalPort       =   713
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ConnectBtn_Click()

    Winsock.Connect
    
End Sub

Private Sub SendBtn_Click()

    Winsock.SendData (SendBox.Text)
    
End Sub

Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)

    Winsock.Accept
    
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim strdata As String
    Winsock.GetData strdata, vbString
    
    RecvBox.Text = strdata
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    MsgBox "Winsock Error"
    
End Sub
