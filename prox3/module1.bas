'---------------------------------------------------------------'
' VB-ASM, Version 1.30                                          '
' Copyright (c) 1994-95 SoftCircuits Programming                '
' Redistributed by Permission.                                  '
'                                                               '
' SoftCircuits Programming                                      '
' P.O. Box 16262                                                '
' Irvine, CA 92713                                              '
' CompuServe: 72134,263                                         '
'                                                               '
' This program may be used and distributed freely on the        '
' condition that it is distributed in full and unchanged, and   '
' that no fee is charged for such use and distribution with the '
' exception or reasonable media and shipping charges.           '
'                                                               '
' You may also incorporate any or all portions of this program, '
' and/or include the VB-ASM DLL, as part of your own programs   '
' and distribute such programs without payment of royalties on  '
' the condition that such program do not duplicate the overall  '
' functionality of VB-ASM and/or any of its demo programs, and  '
' that you agree to the following disclaimer.                   '
'                                                               '
' WARNING: Accessing the low-level services of Windows, DOS and '
' the ROM-BIOS using VB-ASM is an extremely powerful technique  '
' that, if used incorrectly, can cause possible permanent       '
' damage and/or loss of data. You are responsible for           '
' determining appropriate use of any and all files included in  '
' this package. SoftCircuits will not be held liable for any    '
' damages resulting from the use of these files.                '
'                                                               '
' SOFTCIRCUITS SPECIFICALLY DISCLAIMS ALL WARRANTIES,           '
' INCLUDING, WITHOUT LIMITATION, ALL IMPLIED WARRANTIES OF      '
' MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, AND        '
' NON-INFRINGEMENT OF THIRD PARTY RIGHTS.                       '
'                                                               '
' UNDER NO CIRCUMSTANCES WILL SOFTCIRCUITS BE LIABLE FOR        '
' SPECIAL, INCIDENTAL, CONSEQUENTIAL, INDIRECT, OR ANY OTHER    '
' DAMAGES OR CLAIMS ARISING FROM THE USE OF THIS PRODUCT,       '
' INCLUDING LOSS OF PROFITS OR ANY OTHER COMMERCIAL DAMAGES,    '
' EVEN IF WE HAVE BEEN ADVISED OF THE POSSIBILITY OF SUCH       '
' DAMAGES.                                                      '
'                                                               '
' Please contact SoftCircuits Programming if you have any       '
' questions concerning these conditions.                        '
'---------------------------------------------------------------'

'VB-ASM DLL declarations
Type VBregs
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
Declare Sub vbInterrupt Lib "VBASM.DLL" (ByVal IntNum As Integer, InRegs As VBregs, OutRegs As VBregs)
Declare Sub vbinterruptx Lib "VBASM.DLL" (ByVal IntNum As Integer, InRegs As VBregs, OutRegs As VBregs)
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
Declare Function vbRealModeIntX Lib "VBASM.DLL" (ByVal IntNum As Integer, InRegs As VBregs, OutRegs As VBregs) As Integer
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

Global Const maxmenu = 20
Global Const maxmsg = 20
Global Const COMPORT = 0    ' 0 = COM1, 1 = COM2

Type message
    vectors(2, 6) As Integer
    TopLine As String
    BotLine As String
End Type

Type menustate
    choices As Integer
    msg(maxmsg) As message
    choice As Integer
End Type


Global Menus(maxmenu) As menustate

Global SystemState As Integer
Global SystemLine As Integer

Global LastKey As Integer

Sub BacklightOff ()
    
    Dim regs As VBregs

    regs.ax = &H11B         ' backlight off
    regs.dx = COMPORT
    Call vbinterruptx(&H14, regs, regs)
    
    regs.ax = &H100 + Asc("V")
    regs.dx = COMPORT
    Call vbinterruptx(&H14, regs, regs)
    
    regs.ax = &H100 + Asc("@")
    regs.dx = COMPORT
    Call vbinterruptx(&H14, regs, regs)
    
    Form3!BackLightBox.Value = 0

End Sub

Sub BacklightOn ()
    
    Dim regs As VBregs
    
    regs.ax = &H11B         ' backlight on
    regs.dx = COMPORT
    Call vbinterruptx(&H14, regs, regs)
    
    regs.ax = &H100 + Asc("V")
    regs.dx = COMPORT
    Call vbinterruptx(&H14, regs, regs)
    
    regs.ax = &H100 + Asc("A")
    regs.dx = COMPORT
    Call vbinterruptx(&H14, regs, regs)
    
    Form3!BackLightBox.Value = 1


End Sub

Sub ClearRemote ()
    
    Dim regs As VBregs
    
    ' clear remote display with ESC H ESC J
        
    regs.ax = &H11B
    regs.dx = COMPORT
    Call vbinterruptx(&H14, regs, regs)
    
    regs.ax = &H100 + Asc("H")
    regs.dx = COMPORT
    Call vbinterruptx(&H14, regs, regs)
    
    regs.ax = &H11B
    regs.dx = COMPORT
    Call vbinterruptx(&H14, regs, regs)
    
    regs.ax = &H100 + Asc("J")
    regs.dx = COMPORT
    Call vbinterruptx(&H14, regs, regs)
    
End Sub

Sub InitPort ()

    Dim regs As VBregs

    regs.ax = &HE3
    regs.dx = COMPORT
    Call vbinterruptx(&H14, regs, regs)

End Sub

