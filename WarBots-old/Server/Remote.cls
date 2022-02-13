VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim r As String * 42

Public Function getarena() As String

Dim i As Integer
Dim offset As Integer
' r is global for this module (string * 42)

Mid$(r, 1, 1) = Format(LastBot, "0")
Mid$(r, 2, 1) = status

offset = 3

For i = 1 To LastBot
    Mid$(r, offset, 1) = Bots(i).status
    Mid$(r, offset + 1, 3) = Format(Bots(i).x, "000")
    Mid$(r, offset + 4, 3) = Format(Bots(i).y, "000")
    Mid$(r, offset + 7, 3) = Format(Bots(i).newshot, "00")
    Mid$(r, offset + 9, 3) = Format(Bots(i).scan, "0")
    offset = offset + 10
Next i

getarena = r

End Function
Public Function getscan(i As Integer) As String
    
Dim r As String

    r = Format(Bots(i).sdir, "000")
    r = r + Format(Bots(i).sres, "00")
    ' make sure display doesn't see this scan again.
    ' We'll have to change this when we allow two
    ' display processes
    Bots(i).scan = 0
    
    getscan = r

End Function

Public Function getshot(i As Integer) As String
    
Dim r As String

    r = Format(Bots(i).tx, "000")
    r = r + Format(Bots(i).ty, "000")
    If Bots(i).dx < 0 Then
        r = r + Format(Abs(Bots(i).dx), "-00")
    Else
        r = r + Format(Bots(i).dx, "000")
    End If
    If Bots(i).dy < 0 Then
        r = r + Format(Abs(Bots(i).dy), "-00")
    Else
        r = r + Format(Bots(i).dy, "000")
    End If
    r = r + Format(Bots(i).fire, "00")
    
    Bots(i).newshot = 0
    getshot = r

End Function

Private Sub Class_Initialize()

End Sub
