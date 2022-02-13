VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Robot"
Attribute VB_Creatable = True
Attribute VB_Exposed = True
Option Explicit

Public Sub accel(delta As Integer)
   
   Xspeed = Xspeed + delta
   Server.text1 = Xspeed
   
End Sub

Public Sub place(x As Integer, y As Integer)

Server.picture1.Circle (x, y), 50, RGB(255, 0, 0)

End Sub


Property Get Speed() As Integer
Speed = Xspeed
End Property



Private Sub Class_Initialize()

Server.Visible = True

End Sub


