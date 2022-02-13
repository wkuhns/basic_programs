VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2670
   LinkTopic       =   "Form1"
   ScaleHeight     =   1620
   ScaleWidth      =   2670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTimer 
      Caption         =   "&Start Timer"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   900
      Width           =   1935
   End
   Begin VB.CommandButton cmdTestTrigger 
      Caption         =   "&Trigger"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oTrigger As JCTimer.CTrigger
Attribute oTrigger.VB_VarHelpID = -1
Private WithEvents oTimer As JCTimer.CTimer
Attribute oTimer.VB_VarHelpID = -1

Private Sub cmdTestTrigger_Click()
    cmdTestTrigger.Enabled = False
    
    Set oTrigger = New JCTimer.CTrigger
    oTrigger.Start
    
End Sub

Private Sub cmdTimer_Click()
    If oTimer.Enabled Then
        cmdTimer.Caption = "&Start Timer"
    Else
        cmdTimer.Caption = "&Stop Timer"
    End If
    
    oTimer.Enabled = Not oTimer.Enabled
End Sub

Private Sub Form_Load()
    Set oTimer = New JCTimer.CTimer
    oTimer.Interval = 2000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If oTimer.Enabled Then
        oTimer.Enabled = False
    End If
    Set oTimer = Nothing
End Sub

Private Sub oTimer_Timer()
    Beep
End Sub

Private Sub oTrigger_Trigger()
    Set oTrigger = Nothing
    MsgBox "Trigger fired!"
    cmdTestTrigger.Enabled = True
End Sub
