Attribute VB_Name = "ROBOT1"
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

    

Sub ReadJoysticks(a, b, c, d As Integer)
    Dim inregs As vbregs
    
    inregs.ax = &H8400
    inregs.dx = 1
    
    Call vbInterruptX(&H15, inregs, inregs)
    a = inregs.ax
    b = inregs.bx
    c = inregs.cx
    d = inregs.dx
End Sub

