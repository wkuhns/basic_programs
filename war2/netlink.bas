Attribute VB_Name = "netlink"
Option Explicit

Public Sub ParseCommand()

    Dim i As Integer
    Dim buff As String
    Dim unit As Integer

    buff = getcom()

    Select Case Left$(buff, 1)
        Case "U"                            ' get unit info
            unit = Val(Mid$(buff, 3, 3))
            army(1, unit).type = Val(Mid$(buff, 6, 3))
            army(1, unit).index = unit
            army(1, unit).side = THEM
            army(1, unit).speed = Val(Mid$(buff, 10, 6))
            army(1, unit).x = Val(Mid$(buff, 17, 5))
            army(1, unit).y = Val(Mid$(buff, 23, 5))
            army(1, unit).fuel = Val(Mid$(buff, 29, 7))
            army(1, unit).health = Val(Mid$(buff, 37, 5))
            army(1, unit).camo = Val(Mid$(buff, 43, 2))
            'WriteCCC "Unit " + Str$(army(1, unit).Index) + " at " + Str$(army(1, unit).x) + ", " + Str$(army(1, unit).y)
        Case "M"
            terrain(Val(Mid$(buff, 2, 3)), Val(Mid$(buff, 6, 3))).a = Val(Mid$(buff, 10, 4))
        Case "H"
            unit = Val(Mid$(buff, 3, 2))
            army(0, unit).health = army(0, unit).health - Val(Mid$(buff, 6, 6))
            If army(0, unit).health <= 0 Then
                DestroyUnit army(0, unit)
            End If
            SendUnitInfo army(0, unit)
        Case "P"
            MapForm!Timer1.Enabled = False
            MapForm!PauseButton.Enabled = False
            MapForm!RunButton.Enabled = True
        Case "R"
            MapForm!Timer1.Enabled = True
            MapForm!PauseButton.Enabled = True
            MapForm!RunButton.Enabled = False
        Case "D"
            MapForm.MessageBox.Text = "Getting Map"
            DisplayMap
            MapForm.MessageBox.Text = "Got Map"
            SetupPlayer
            CCCForm.Show
        Case Else
            MsgBox "Bad Buffer: " & buff
    End Select

End Sub
Sub SendUnitInfo(unit As unitstruct)

    Static buff As String

    buff = "U " + Format$(unit.index, "00") + " "
    buff = buff + Format$(unit.type, "000") + " "
    buff = buff + Format$(unit.speed, "000.00") + " "
    buff = buff + Format$(unit.x, "00.00") + " "
    buff = buff + Format$(unit.y, "00.00") + " "
    buff = buff + Format$(unit.fuel, "0000.00") + " "
    buff = buff + Format$(unit.health, "00.00") + " "
    buff = buff + Format$(unit.camo, "00")

    SendCom buff

End Sub

Sub SendHitInfo(unit As unitstruct, impact As Single)

    Static buff As String

    buff = "H " + Format$(unit.index, "00") + " "
    buff = buff + Format$(impact, "000.00")

    SendCom buff

End Sub
