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
   Begin VB.Timer Timer2 
      Left            =   968
      Top             =   242
   End
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

MyForm.Timer2.Enabled = False
Set Mybot = Nothing

End

End Sub


'
' Don't change this at all. This code creates the linkage
' to the robot server process.
'
Private Sub Form_Load()

' Create robot object

Set Mybot = CreateObject("RobotAPI.RobotLink")

' Register 'Ping' procedure with server.

Call Mybot.RegisterAlert(MyForm)

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
    If Mybot.Status = "K" Then
        Die
        Exit Sub
    End If
    
    ' Do the user's cyclic stuff.
    If Mybot.server_status = "R" Then
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
    e = Mybot.pinged
    If e <> 0 Then
        Ping (e)
    End If

End Sub
Private Sub Timer2_Timer()

Dim delta As Integer

delta = newdir - dir

If delta <> 0 Then
    If delta > 180 Then delta = delta - 360
    If delta < -180 Then delta = delta + 360

    If delta > 9 Then delta = 9

    If delta < -9 Then delta = -9

    dir = (dir + delta + 360) Mod 360
Else
    Timer2.Enabled = False
End If

End Sub


