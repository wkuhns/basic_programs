VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Private Sub Class_Initialize()

Set arena = New ArenaForm
Set serverbox = New serverform

arena.Enabled = True
arena.Visible = True
serverbox.Enabled = True
serverbox.Visible = True

    Randomize
    
    Bots(1).color = RGB(255, 0, 0)
    Bots(2).color = RGB(0, 255, 0)
    Bots(3).color = RGB(0, 0, 255)
    Bots(4).color = RGB(255, 0, 255)
    quads(1).x = 100
    quads(1).y = 100
    quads(2).x = 100
    quads(2).y = 600
    quads(3).x = 600
    quads(3).y = 100
    quads(4).x = 600
    quads(4).y = 600

    DebugState = False

End Sub


Private Sub Class_Terminate()

Set arena = Nothing
Set serverbox = Nothing

End Sub
