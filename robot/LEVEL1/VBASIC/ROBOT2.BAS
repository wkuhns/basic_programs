Attribute VB_Name = "ROBOT21"
'VB-ASM DLL declarations
Type vbregs
    ax As Integer
    bx As Integer
    cx As Integer
    dx As Integer
    BP As Integer
    SI As Integer
    DI As Integer
    Flags As Integer
    DS As Integer
    ES As Integer
End Type

Declare Sub vbInterruptX Lib "VBASM.DLL" (ByVal IntNum As Integer, inregs As vbregs, OutRegs As vbregs)

Type joint
    name As String
    MinAngle As Integer
    MaxAngle As Integer
    CalArray(2, 5) As Single
    NegDirection As Integer
    Posdirection As Integer
    MinSpeed As Integer
    MaxSpeed As Integer
    DecelRange As Integer
    TargetAngle As Single
    CurrentSpeed As Integer
    CurrentDirection As Integer
    MotorFlag As Integer
    Counts As Integer
    ActiveFlag As Integer
    s  As Integer
    x As Single
    y As Single
    z As Single
    r As Single
    rz As Single
    ry As Single
    theta As Single
    nt As Single
End Type

Global j(4) As joint

Sub GetDegrees(ch As Integer)
    If j(ch).CalArray(2, 1) > j(ch).CalArray(2, 2) Then
                       ' descending table
        i = 5
        While (j(ch).CalArray(2, i) < j(ch).s) And (i > 1)
            i = i - 1
        Wend

        a1 = j(ch).CalArray(1, i)
        s1 = j(ch).CalArray(2, i)
        a2 = j(ch).CalArray(1, i + 1)
        s2 = j(ch).CalArray(2, i + 1)
        j(ch).theta = (((s1 - j(ch).s) * (a2 - a1)) / (s1 - s2)) + a1
    Else

        i = 5
        While (j(ch).CalArray(2, i) < j(ch).s) And (i < 5)
            i = i + 1
        Wend

        a1 = j(ch).CalArray(1, i - 1)
        s1 = j(ch).CalArray(2, i - 1)
        a2 = j(ch).CalArray(1, i)
        s2 = j(ch).CalArray(2, i)
        j(ch).theta = (((j(ch).s - s1) * (a2 - a1)) / (s2 - s1)) + a1
        
    End If
End Sub

Sub ReadCalFile(channel As Integer)
    
    filename$ = "d:\ROBOT\LEVEL1\CAL" + Mid$(Str(channel), 2, 1) + ".DAT"
    
    Open filename$ For Input As #1
    For i = 1 To 5
        Input #1, j(channel).CalArray(1, i), j(channel).CalArray(2, i)
    Next i
    Close #1

End Sub


Sub ReadJoysticks()
    Dim inregs As vbregs
    
    inregs.ax = &H8400
    inregs.dx = 1
    
    Call vbInterruptX(&H15, inregs, inregs)
    j(1).s = inregs.ax
    j(2).s = inregs.bx
    j(3).s = inregs.cx
    j(4).s = inregs.dx
End Sub

Sub resolve()
    
    theta1 = (j(1).theta / 57.296)
    theta2 = (j(2).theta / 57.296) - 40

    j(2).x = j(1).x + j(1).r * Sin(theta1)
    j(2).y = j(1).y + j(1).r * Cos(theta1)
    j(2).z = j(1).z
    
    ' j(2) axis is parallel to z plane. calculate x projection
    j(2).rz = j(2).r * Cos(theta2)

    j(3).x = j(2).x + j(2).rz * Sin(theta1)
    j(3).y = j(2).y + j(2).rz * Cos(theta1)
    j(3).z = j(2).z + (j(2).r * Sin(theta2))

    ' nt is Net Theta
    j(3).nt = j(3).theta + j(2).theta
    theta3 = 120 - (j(3).nt / 57.296)
    
    j(3).rz = j(3).r * Cos(theta3)

    j(4).x = j(3).x + (j(3).rz * Sin(theta1))
    j(4).y = j(3).y + (j(3).rz * Cos(theta1))
    j(4).z = j(3).z + (j(3).r * Sin(theta3))

End Sub

Sub WriteCalFile(channel As Integer)
    
    filename$ = "d:\robot\level1\CAL" + Mid$(Str(channel), 2, 1) + ".DAT"
    
    Open filename$ For Output As #1
    For i = 1 To 5
        Print #1, j(channel).CalArray(1, i), j(channel).CalArray(2, i)
    Next i
    Close #1

End Sub

