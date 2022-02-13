'VB-ASM DLL declarations
Type vbregs
    ax As Integer
    BX As Integer
    CX As Integer
    dx As Integer
    BP As Integer
    SI As Integer
    DI As Integer
    Flags As Integer
    DS As Integer
    ES As Integer
End Type

'VBREGS Flags bit values
Global Const VBFLAGS_CARRY = &H1
Global Const VBFLAGS_PARITY = &H4
Global Const VBFLAGS_AUX = &H10
Global Const VBFLAGS_ZERO = &H40
Global Const VBFLAGS_SIGN = &H80

'vbGetDriveType return values
Global Const VBDRIVE_REMOVABLE = &H1002
Global Const VBDRIVE_FIXED = &H1003
Global Const VBDRIVE_REMOTE = &H1004
Global Const VBDRIVE_CDROM = &H1005
Global Const VBDRIVE_FLOPPY = &H1006
Global Const VBDRIVE_RAMDISK = &H1007
Global Const VBDRIVE_UNKNOWN = &H1010
Global Const VBDRIVE_INVALID = &H0

Declare Function vbGetCtrlHwnd Lib "VBASM.DLL" (Ctrl As Any) As Integer
Declare Function vbGetCtrlModel Lib "VBASM.DLL" (Ctrl As Any) As Long
Declare Function vbGetCtrlName Lib "VBASM.DLL" (Ctrl As Any) As String
Declare Sub vbGetData Lib "VBASM.DLL" (ByVal Pointer As Long, Variable As Any, ByVal nCount As Integer)
Declare Function vbGetDriveType Lib "VBASM.DLL" (ByVal drive As String) As Integer
Declare Function vbGetLongPtr Lib "VBASM.DLL" (nVariable As Any) As Long
Declare Function vbHiByte Lib "VBASM.DLL" (ByVal nValue As Integer) As Integer
Declare Function vbHiWord Lib "VBASM.DLL" (ByVal nValue As Long) As Integer
Declare Function vbInp Lib "VBASM.DLL" (ByVal nPort As Integer) As Integer
Declare Function vbInpw Lib "VBASM.DLL" (ByVal nPort As Integer) As Integer
Declare Sub vbInterrupt Lib "VBASM.DLL" (ByVal IntNum As Integer, InRegs As vbregs, OutRegs As vbregs)
Declare Sub vbinterruptx Lib "VBASM.DLL" (ByVal IntNum As Integer, InRegs As vbregs, OutRegs As vbregs)
Declare Function vbLoByte Lib "VBASM.DLL" (ByVal nValue As Integer) As Integer
Declare Function vbLoWord Lib "VBASM.DLL" (ByVal nValue As Long) As Integer
Declare Function vbMakeLong Lib "VBASM.DLL" (ByVal nLoWord As Integer, ByVal nHiWord As Integer) As Long
Declare Function vbMakeWord Lib "VBASM.DLL" (ByVal nLoByte As Integer, ByVal nHiByte As Integer) As Integer
Declare Sub vbout Lib "VBASM.DLL" (ByVal nPort As Integer, ByVal nData As Integer)
Declare Sub vbOutw Lib "VBASM.DLL" (ByVal nPort As Integer, ByVal nData As Integer)
Declare Function vbPeek Lib "VBASM.DLL" (ByVal nSegment As Integer, ByVal nOffset As Integer) As Integer
Declare Function vbPeekw Lib "VBASM.DLL" (ByVal nSegment As Integer, ByVal nOffset As Integer) As Integer
Declare Sub vbPoke Lib "VBASM.DLL" (ByVal nSegment As Integer, ByVal nOffset As Integer, ByVal nValue As Integer)
Declare Sub vbPokew Lib "VBASM.DLL" (ByVal nSegment As Integer, ByVal nOffset As Integer, ByVal nValue As Integer)
Declare Function vbRealModeIntX Lib "VBASM.DLL" (ByVal IntNum As Integer, InRegs As vbregs, OutRegs As vbregs) As Integer
Declare Function vbRecreateCtrl Lib "VBASM.DLL" (Ctrl As Any) As Integer
Declare Function vbSAdd Lib "VBASM.DLL" (Variable As String) As Integer
Declare Sub vbSetData Lib "VBASM.DLL" (ByVal Pointer As Long, Variable As Any, ByVal nCount As Integer)
Declare Function vbShiftLeft Lib "VBASM.DLL" (ByVal nValue As Integer, ByVal nBits As Integer) As Integer
Declare Function vbShiftLeftLong Lib "VBASM.DLL" (ByVal nValue As Long, ByVal nBits As Integer) As Long
Declare Function vbShiftRight Lib "VBASM.DLL" (ByVal nValue As Integer, ByVal nBits As Integer) As Integer
Declare Function vbShiftRightLong Lib "VBASM.DLL" (ByVal nValue As Long, ByVal nBits As Integer) As Long
Declare Function vbSSeg Lib "VBASM.DLL" (Variable As String) As Integer
Declare Function vbVarPtr Lib "VBASM.DLL" (Variable As Any) As Integer
Declare Function vbVarSeg Lib "VBASM.DLL" (Variable As Any) As Integer

Global Const COM1 = 0
Global Const COM2 = 1

Global comport As Integer

Global CRLF As String

Sub getcom (buff As String)
    
    Dim regs As vbregs
    Dim i As Long

    buff = ""
    i = 1
    
    regs.ax = &H300
    regs.dx = comport
    Call vbinterruptx(&H14, regs, regs)
    While (regs.ax And &H100) = 0
        regs.ax = &H300
        regs.dx = comport
        Call vbinterruptx(&H14, regs, regs)
    Wend
    
    regs.ax = &H200
    regs.dx = comport
    Call vbinterruptx(&H14, regs, regs)
    
    While Abs(regs.ax) <= 127 And regs.ax <> 13
        buff = buff + Chr$(regs.ax)
        i = i + 1
        regs.ax = &H200
        regs.dx = comport
        Call vbinterruptx(&H14, regs, regs)
    Wend

    regs.ax = &H200
    regs.dx = comport
    Call vbinterruptx(&H14, regs, regs)


End Sub

Sub InitComPort ()

    Dim regs As vbregs

    regs.ax = &HE3
    regs.dx = comport
    Call vbinterruptx(&H14, regs, regs)
    CRLF = Chr$(13) + Chr$(10)

End Sub

Sub sendcom (buff As String)
    
    Dim regs As vbregs
    Dim i As Long

    For i = 1 To Len(buff)
        regs.ax = &H100 + Asc(Mid$(buff, i, 1))
        regs.dx = comport
        Call vbinterruptx(&H14, regs, regs)
    Next i

End Sub

