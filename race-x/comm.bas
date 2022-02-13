Option Explicit


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

Type ComStat
    Status      As String * 1
    inqueue     As Integer
    outqueue    As Integer
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
Declare Function GetCommError Lib "user" (ByVal a As Integer, b As ComStat) As Integer

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

Global comstatus As ComStat

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

Global apierrnum           As Integer

Sub InitializePort ()

    Dim apierr As Integer

    CommHandle = -1
    CommDeviceNum = -1
    
    ' Default Port Settings

    CommPortName = "COM1:"
    CommState.BaudRate = 9600
    CommState.ByteSize = Chr$(7)
    CommState.Parity = "n"
    CommState.StopBits = Chr$(1)
    
    ' Default Line Settings

    CommRBBuffer = 2048
    CommTBBuffer = 2048
    CommState.RlsTimeOut = 0
    CommState.CtsTimeOut = 0
    CommState.DsrTimeOut = 0
    CommEventMask = &H3FF
    CommReadInterval = 500

    CommHandle = OpenComm(CommPortName, CommRBBuffer, CommTBBuffer)

    If CommHandle = -2 Then
        Debug.Print "Port already Open!"
        Debug.Print "Attempting restart"
        
        apierrnum = CloseComm(0)
        CommHandle = OpenComm(CommPortName, CommRBBuffer, CommTBBuffer)

    End If

    If CommHandle < 0 Then
        Debug.Print " OpenComm() API Failed! (ERR "; Str$(CommHandle) + ")"
    Else
        apierr = SetCommEventMask(CommHandle, CommEventMask)

        CommState.Id = Chr$(CommHandle)

        apierr = SetCommState(CommState)

    End If

End Sub

Function ReadCommPort (ReadAmount As Integer) As String
    
    Dim apierr As Integer
    Dim EventMask As Integer
    Dim Found As Integer
    Dim buffer$

    If ReadAmount < 1 Then
        ReadCommPort = ""
        Exit Function
    End If

    EventMask = CommEventMask
    apierr = GetCommEventMask(CommHandle, EventMask)
    
    If apierr And EV_RXCHAR Then
        buffer$ = Space$(ReadAmount)
        apierr = ReadComm(CommHandle, buffer$, Len(buffer$))

        If apierr <= 0 Then
            Debug.Print " ReadCOMM API FAILED! (ERR "; Str$(apierr); ")"
            buffer$ = ""
            apierr = GetCommError(CommHandle, comstatus)
        Else
            buffer$ = Left$(buffer$, apierr)
        End If
    End If

'    If (ApiErr And EV_RXFLAG) And (CommEventMask And EV_RXFLAG) Then
'    End If

'    If (ApiErr And EV_TXEMPTY) And (CommEventMask And EV_XFLAG) Then
'    End If

'    If (ApiErr And EV_CTS) And (CommEventMask And EV_CTS) Then
'    End If

'    If (ApiErr And EV_DSR) And (CommEventMask And EV_DSR) Then
'    End If

'    If (ApiErr And EV_RLSD) And (CommEventMask And EV_RLSD) Then
'    End If

'    If (ApiErr And EV_BREAK) And (CommEventMask And EV_BREAK) Then
'    End If

'    If (ApiErr And EV_ERR) And (CommEventMask And EV_ERR) Then
'    End If
    
'    If (ApiErr And EV_PERR) And (CommEventMask And EV_PERR) Then
'    End If
    
'    If (ApiErr And EV_RING) And (CommEventMask And EV_RING) Then
'        UpdateCaption " Receive Window: RING! ", 0
'        Beep
'    End If
    
    ReadCommPort = buffer$

End Function

Sub WriteCommPort (Send$)

    Dim apierr%

    apierr% = WriteComm(CommHandle, Send$, Len(Send$))

    If apierr% < 0 Then
        Debug.Print "WriteComm API Failed! (ERR "; Str$(apierr%); ")"
    End If

End Sub

