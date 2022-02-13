
' Communication Demo is a sample aplication showing how the
' Windows COMM API function can be used in a Visual Basic program.
' This sample program does not utilize all the functions available
' through the Windows COMM API function, but is a can be used as
' a starting point.                                                ==

Type CommStateDCB
    Id          As String * 1   ' Port Id from OpenComm
    BaudRate    As Integer      ' Baud Rate
    ByteSize    As String * 1   ' Data Bit Size (4 to 8)
    Parity      As String * 1   ' Parity
    StopBits    As String * 1   ' Stop Bits
    RlsTimeOut  As Integer      ' Carrier Detect Time "CD"
    CtsTimeOut  As Integer      ' Clear-to-Send Time
    DsrTimeOut  As Integer      ' Data-Set-Ready Time
    ModeControl As Integer      ' Mode Control Bit Fields
    XonChar     As String * 1   ' XON character
    XoffChar    As String * 1   ' XOFF character
    XonLim      As Integer      ' Min characters in buffer before XON is sent
    XoffLim     As Integer      ' Max characters in buffer before XOFF is send
    PeChar      As String * 1   ' Parity Error Character
    EofChar     As String * 1   ' EOF/EOD character
    EvtChar     As String * 1   ' Event character
    TxDelay     As Integer      ' Reserved/Not Used
End Type

Type TextMetrics
    tmHeight As Integer
    work1 As String * 14
    work2 As String * 9
    Work3 As String * 6
End Type

Declare Function OpenComm Lib "user" (ByVal a As String, ByVal b As Integer, ByVal c As Integer) As Integer
Declare Function CloseComm Lib "user" (ByVal a As Integer) As Integer

Declare Function WriteComm Lib "user" (ByVal a As Integer, ByVal b As String, ByVal c As Integer) As Integer
Declare Function ReadComm Lib "user" (ByVal a As Integer, ByVal b As String, ByVal c As Integer) As Integer

Declare Function GetCommEventMask Lib "user" (ByVal a As Integer, ByVal b As Integer) As Integer
Declare Function SetCommEventMask Lib "user" (ByVal a As Integer, ByVal b As Integer) As Integer

Declare Function SetCommState Lib "user" (b As CommStateDCB) As Integer
Declare Function GetCommState Lib "user" (ByVal a As Integer, b As CommStateDCB) As Integer

Declare Function RemoveMenu Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Declare Function GetSystemMenu Lib "User" (ByVal hWnd As Integer, ByVal Action As Integer) As Integer

'Global Const FALSE = 0
'Global Const TRUE = Not FALSE

Global Const MF_BYPOSITION = &H400  'Used by RemoveMenu()

' COMM OPEN Error Numbers

Global Const IE_BADID = -1      ' Invalid or unsupported id
Global Const IE_OPEN = -2       ' Device Already Open
Global Const IE_NOPEN = -3      ' Device Not Open
Global Const IE_MEMORY = -4     ' Unable to allocate queues
Global Const IE_DEFAULT = -5    ' Error in default parameters
Global Const IE_HARDWARE = -10  ' Hardware Not Present
Global Const IE_BYTESIZE = -11  ' Illegal Byte Size
Global Const IE_BAUDRATE = -12  ' Unsupported BaudRate

' COMM EVENT MASK

Global Const EV_RXCHAR = &H1
Global Const EV_RXFLAG = &H2
Global Const EV_TXEMPTY = &H4
Global Const EV_CTS = &H8
Global Const EV_DSR = &H10
Global Const EV_RLSD = &H20
Global Const EV_BREAK = &H40
Global Const EV_ERR = &H80
Global Const EV_RING = &H100
Global Const EV_PERR = &H200
Global Const EV_ALL = &H3FF
            
Global CommHandle As Integer
Global CommDeviceNum As Integer

Global CommPortName As String
Global PostPortName As String

Global CommEventMask As Integer
Global PostEventMask As Integer

Global CommState As CommStateDCB
Global PostState As CommStateDCB

Global CommRBBuffer As Integer
Global PostRBBuffer As Integer

Global CommTBBuffer As Integer
Global PostTBBuffer As Integer

Global CommReadInterval As Integer
Global PostReadInterval As Integer

Global CaptionLeft      As Integer
Global CaptionTop       As Integer
Global CaptionHeight    As Integer
Global CaptionCenter    As Integer
Global CaptionWidth     As Integer
Global CaptionText$


