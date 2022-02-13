VERSION 5.00
Begin VB.Form MyForm 
   Caption         =   "Form1"
   ClientHeight    =   605
   ClientLeft      =   1727
   ClientTop       =   2244
   ClientWidth     =   1518
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   605
   ScaleWidth      =   1518
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   242
      Top             =   242
   End
End
Attribute VB_Name = "MyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Our life is over. Release robot's soul and die. Do not
' change this subroutine
'
Sub Die()

Set MyBot = Nothing

End

End Sub


'
' Don't change this at all. This code creates the linkage
' to the robot server process.
'
Private Sub Form_Load()

' Create robot object

Set MyBot = CreateObject("RobotAPI.RobotLink")

' Register 'Ping' procedure with server.

Call MyBot.RegisterAlert(MyForm)

Timer1.Interval = 100
Timer1.Enabled = True
' Do user's initialization.

UserInit

' Don't change this - User specific stuff is in DoFrame.

While True
    ' Check to see if we're dead. You can't cheat death
    ' by changing this - all that will happen is that
    ' you'll have dead processes cluttering up your
    ' system.
    If MyBot.Status = "K" Then
        Die
        Exit Sub
    End If
    
    ' Do the user's cyclic stuff.
    If MyBot.server_status = "R" Then
        UserFrame
    Else
        DoEvents
    End If
    
Wend

End Sub

Private Sub Form_Unload(Cancel As Integer)

MsgBox ("AppForm unloading...")

Die

End Sub

Private Sub Timer1_Timer()

Dim e As Integer
    e = MyBot.pinged
    If e <> 0 Then
        Ping (e)
    End If

End Sub
