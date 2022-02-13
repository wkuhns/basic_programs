'------------------------------------------------------------------------------
'
'       This application provides an on-screen window with the functions of the
'       Proximity BITE remote unit. It also provides serial communications to an
'       actual remote unit.
'
'       A data file named prox.cfg is read to get text for error messages and
'       text for channel descriptions.
'
'       When the ON/OFF button is pressed, a popup window allows selection of a
'       menu data file.
'
'------------------------------------------------------------------------------

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

'VBASM.DLL function declarations

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

'------------------------------------------------------------------------------
'   Global constants
'------------------------------------------------------------------------------

Global Const MAXMENU = 30
Global Const MAXMSG = 20
Global Const NUMCHANNELS = 8            ' 8 prox channels
Global Const THRESHOLD = .25            ' prox near/far threshold

Global Const COMPORT = 0                ' 0 = COM1, 1 = COM2
Global Const REMOTE = 0                 ' 1 = use remote, 0 = no comm

Global Const TOP_MENU = 1
Global Const SYSTEM_OUTPUTS_MENU = 2
Global Const PRESENT_STATUS_MENU = 3
Global Const PRESENT_FAULT_MENU = 4
Global Const FAULT_BY_LEG_MENU = 5
Global Const FBL1_MENU = 6
Global Const FBL2_MENU = 7
Global Const FBL3_MENU = 8
Global Const FBL4_MENU = 9
Global Const FBL32_MENU = 10
Global Const SELF_TEST_MENU = 11
Global Const FAULT_BY_CHANNEL_MENU = 12
Global Const LMG_FAULT_MENU = 13
Global Const LGUL_FAULT_MENU = 14
Global Const LDOOR_FAULT_MENU = 15
Global Const LGDL_FAULT_MENU = 16
Global Const NGUL_FAULT_MENU = 17
Global Const MONITOR_MENU = 18
Global Const NG_FAULT_MENU = 19
Global Const HIDDEN_MENU = 20
Global Const MONITOR1_MENU = 21
Global Const MONITOR2_MENU = 22
Global Const MONITOR3_MENU = 23

Global Const LGUL = 0
Global Const LGDL = 1
Global Const LDOOR = 2
Global Const RGUL = 3
Global Const RGDL = 4
Global Const RDOOR = 5
Global Const NGUL = 6
Global Const NGDL = 7

'------------------------------------------------------------------------------
'    Structures
'------------------------------------------------------------------------------

Type menupage
    vectors(2, 6) As Integer            ' Keypad vectors: menu,line
    line1 As String                     ' Display lines 1-4
    line2 As String
    line3 As String
    line4 As String
End Type

Type menustruct                         ' Structure for a menu
    pagecount           As Integer      ' Size of this menu
    page(MAXMSG)        As menupage     ' array of menu lines
    choice              As Integer      ' current line
End Type

Type errorstruct                        ' BITE error structure
    errcount    As Integer
    channel     As Integer
    errnum      As Integer              ' index to errtext array
    leg         As Integer              ' flight leg
    gap         As Single               ' measured gap
End Type

Type channelstruct                      ' structure for a channel
    name        As String
    abbr        As String               ' abbreviation
    ntext       As String               ' text for near
    ftext       As String               ' text for far
    errnum      As Integer              ' error, if any
    gap         As Single
    mingap      As Single
    nomgap      As Single
    maxgap      As Single
End Type

'------------------------------------------------------------------------------
'   Global variables and arrays
'------------------------------------------------------------------------------

Global menus(MAXMENU)           As menustruct
Global Channels(NUMCHANNELS)    As channelstruct
Global BiteMsgs(50)             As errorstruct
Global SystemMenu               As Integer
Global SystemPage               As Integer
Global ErrText(20)              As String

Global keybuffer                As String

Global last_leg                                 As Integer

'------------------------------------------------------------------------------
'       Turn off QTERM backlight
'------------------------------------------------------------------------------
'
Sub BacklightOff ()
    
    Dim regs As VBregs

	If REMOTE Then
	    regs.ax = &H11B                                                             ' backlight off
	    regs.dx = COMPORT
	    Call vbinterruptx(&H14, regs, regs)
    
	    regs.ax = &H100 + Asc("V")
	    regs.dx = COMPORT
	    Call vbinterruptx(&H14, regs, regs)
    
	    regs.ax = &H100 + Asc("@")
	    regs.dx = COMPORT
	    Call vbinterruptx(&H14, regs, regs)
	End If

    Form3!BackLightBox.Value = 0

End Sub

'------------------------------------------------------------------------------
'       Turn on QTERM backlight
'------------------------------------------------------------------------------
'
Sub BacklightOn ()
    
    Dim regs As VBregs
    
	If REMOTE Then
	    regs.ax = &H11B         ' backlight on
	    regs.dx = COMPORT
	    Call vbinterruptx(&H14, regs, regs)
    
	    regs.ax = &H100 + Asc("V")
	    regs.dx = COMPORT
	    Call vbinterruptx(&H14, regs, regs)
    
	    regs.ax = &H100 + Asc("A")
	    regs.dx = COMPORT
	    Call vbinterruptx(&H14, regs, regs)
	End If
    
    Form3!BackLightBox.Value = 1
	

End Sub

'------------------------------------------------------------------------------
'       Clear QTERM display
'------------------------------------------------------------------------------
'
Sub ClearRemote ()
    
    Dim regs As VBregs
    
	If REMOTE Then
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
	End If

End Sub

'------------------------------------------------------------------------------
'       Display all four lines of current menu page
'------------------------------------------------------------------------------
'
Sub DisplayPage ()
	
    Static line1, line2, line3, line4 As String

    m = SystemMenu
    l = SystemPage

	' Only display if new / changed data

    If (line1 <> menus(m).page(l).line1 Or line2 <> menus(m).page(l).line2 Or line3 <> menus(m).page(l).line3 Or line4 <> menus(m).page(l).line4) Then
	
		line1 = menus(m).page(l).line1
		line2 = menus(m).page(l).line2
		line3 = menus(m).page(l).line3
		line4 = menus(m).page(l).line4
	
		' paint new text to local window

		Form1.Line1.Text = menus(m).page(l).line1
		Form1.Line2.Text = menus(m).page(l).line2
		Form1.Line3.Text = menus(m).page(l).line3
		Form1.Line4.Text = menus(m).page(l).line4
	
		' clear remote display with ESC H ESC J
	
		ClearRemote

		' send lines (with CR) to remote display

		sendchars (menus(m).page(l).line1)
	    sendCR

		sendchars (menus(m).page(l).line2)
		sendCR
    
		sendchars (menus(m).page(l).line3)
		sendCR

		sendchars (menus(m).page(l).line4)
	
		' show options on debug display
    
		Dform!Text1.Text = "Current: " + Str(SystemMenu) + "," + Str(SystemPage)
		Dform!Text2.Text = "Yes: " + Str(menus(SystemMenu).page(SystemPage).vectors(0, 1)) + "," + Str(menus(SystemMenu).page(SystemPage).vectors(1, 1))
		Dform!Text3.Text = "Up: " + Str(menus(SystemMenu).page(SystemPage).vectors(0, 2)) + "," + Str(menus(SystemMenu).page(SystemPage).vectors(1, 2))
		Dform!Text4.Text = "On: " + Str(menus(SystemMenu).page(SystemPage).vectors(0, 3)) + "," + Str(menus(SystemMenu).page(SystemPage).vectors(1, 3))
		Dform!Text5.Text = "No: " + Str(menus(SystemMenu).page(SystemPage).vectors(0, 4)) + "," + Str(menus(SystemMenu).page(SystemPage).vectors(1, 4))
		Dform!Text6.Text = "Down: " + Str(menus(SystemMenu).page(SystemPage).vectors(0, 5)) + "," + Str(menus(SystemMenu).page(SystemPage).vectors(1, 5))
		Dform!Text7.Text = "Menu: " + Str(menus(SystemMenu).page(SystemPage).vectors(0, 6)) + "," + Str(menus(SystemMenu).page(SystemPage).vectors(1, 6))
    End If

End Sub

'------------------------------------------------------------------------------
'       When gap is changed, update channel structure and display widget
'------------------------------------------------------------------------------
'
Sub GapChange (index As Integer)
    
    Channels(index).gap = SensorForm.Gap1Bar(index).Value / 100
    SensorForm.Text1(index).Text = Str$(SensorForm.Gap1Bar(index) / 100)

End Sub

'------------------------------------------------------------------------------
'       Initialize COM port: 9600/8/1/n
'------------------------------------------------------------------------------
'
Sub InitPort ()

    Dim regs As VBregs

	If REMOTE Then
	    regs.ax = &HE3
	    regs.dx = COMPORT
	    Call vbinterruptx(&H14, regs, regs)
	End If

End Sub

'------------------------------------------------------------------------------
'       create unchanging parts of Present Status menu
'       There are two menu pages per channel
'------------------------------------------------------------------------------
'
Sub InitPSmenu ()

    Dim i, m, l As Integer
    Dim Text1, Text2 As String

    m = PRESENT_STATUS_MENU
    l = 0                               ' first line

    menus(m).pagecount = NUMCHANNELS * 2

    For i = 0 To NUMCHANNELS - 1
		Text1 = Channels(i).abbr
		Text2 = Str$(l + 1) + "/" + Format$(2 * NUMCHANNELS)

		menus(m).page(l).line1 = Text1 + Space$(20 - Len(Text1) - Len(Text2)) + Text2

		l = l + 1                       ' Make next page

		Text1 = Channels(i).abbr
		Text2 = Str$(l + 1) + "/" + Format$(2 * NUMCHANNELS)
		menus(m).page(l).line1 = Text1 + Space$(20 - Len(Text1) - Len(Text2)) + Text2

		l = l + 1                       ' Make next page
	Next i

	MakeCircular (m)
	Call SetParent(m, TOP_MENU, 1)
    Call MakeChildMenu(m, HIDDEN_MENU, 1)

End Sub

'------------------------------------------------------------------------------
'   create unchanging parts of monitor  menus. All three are children of
'       toplevel menu #7
'------------------------------------------------------------------------------
'
Sub InitRTmenu ()

    Dim m As Integer
    Dim Text1, Text2 As String

    m = MONITOR_MENU

	'-------------
	'MONITOR_MENU is a two page menu that asks the user which group of channels
	'to display.
	'-------------

    menus(m).pagecount = 3

    menus(m).page(0).line1 = "MONITOR DATA FOR 1/3"
    menus(m).page(0).line2 = "LEFT MAIN GEAR?     "
    menus(m).page(0).vectors(0, 1) = MONITOR1_MENU
    menus(m).page(0).vectors(1, 1) = 0
    menus(m).page(0).vectors(0, 4) = m
    menus(m).page(0).vectors(1, 4) = 1
    
    menus(m).page(1).line1 = "MONITOR DATA FOR 2/3"
    menus(m).page(1).line2 = "RIGHT MAIN GEAR?    "
    menus(m).page(1).vectors(0, 1) = MONITOR2_MENU
    menus(m).page(1).vectors(1, 1) = 0
    menus(m).page(1).vectors(0, 4) = m
    menus(m).page(1).vectors(1, 4) = 2

    menus(m).page(2).line1 = "MONITOR DATA FOR 3/3"
    menus(m).page(2).line2 = "NOSE GEAR?          "
    menus(m).page(2).vectors(0, 1) = MONITOR3_MENU
    menus(m).page(2).vectors(1, 1) = 0
    menus(m).page(2).vectors(0, 4) = m
    menus(m).page(2).vectors(1, 4) = 0

    MakeCircular (m)
    Call SetParent(m, TOP_MENU, 3)

	'-------------
	'MONITOR1_MENU is a three page menu, with a page for each left main sensors
	'-------------

    m = MONITOR1_MENU
    menus(m).pagecount = 3

    For i = 0 To 2
	Text1 = "CH " + Format$(i + 1) + " (" + Channels(i).abbr + ")"
	Text2 = Str$(l + 1) + "/3"
	menus(m).page(i).line1 = Text1 + Space$(20 - Len(Text1) - Len(Text2)) + Text2
	menus(m).page(i).line2 = Channels(i).name
    Next i

    MakeCircular (m)
    Call SetParent(m, MONITOR_MENU, 0)

	'-------------
	'MONITOR2_MENU is a three page menu, with a page for each right main sensor
	'-------------

    m = MONITOR2_MENU
    menus(m).pagecount = 3

    For i = 3 To 5
	Text1 = "CH " + Format$(i + 1) + " (" + Channels(i).abbr + ")"
	Text2 = Str$(l + 1) + "/3"
	menus(m).page(i - 3).line1 = Text1 + Space$(20 - Len(Text1) - Len(Text2)) + Text2
	menus(m).page(i - 3).line2 = Channels(i).name
    Next i

    MakeCircular (m)
    Call SetParent(m, MONITOR_MENU, 1)

	'-------------
	'MONITOR3_MENU is a two page menu, with a page for each nose gear sensor
	'-------------

    m = MONITOR3_MENU
    menus(m).pagecount = 2

    For i = 6 To 7
	Text1 = "CH " + Format$(i + 1) + " (" + Channels(i).abbr + ")"
	Text2 = Str$(l + 1) + "/2"
	menus(m).page(i - 6).line1 = Text1 + Space$(20 - Len(Text1) - Len(Text2)) + Text2
	menus(m).page(i - 6).line2 = Channels(i).name
    Next i

    MakeCircular (m)
    Call SetParent(m, MONITOR_MENU, 2)

End Sub

'------------------------------------------------------------------------------
'       create unchanging parts of System Outputs menu
'------------------------------------------------------------------------------
'
Sub InitSOmenu ()

    Dim m As Integer
    Dim Text1, Text2 As String

    m = SYSTEM_OUTPUTS_MENU

    menus(m).pagecount = 7

	menus(m).page(0).line1 = "LEFT GEAR SAFE   1/7"
	menus(m).page(1).line1 = "LEFT GEAR UNSAFE 2/7"
	menus(m).page(2).line1 = "RT GEAR SAFE     3/7"
	menus(m).page(3).line1 = "RT GEAR UNSAFE   4/7"
	menus(m).page(4).line1 = "NOSE GEAR SAFE   5/7"
	menus(m).page(5).line1 = "NOSE GEAR UNSAFE 6/7"
	menus(m).page(6).line1 = "GEAR DOORS       7/7"

	MakeCircular (m)
	Call SetParent(m, TOP_MENU, 4)
    Call MakeChildMenu(m, HIDDEN_MENU, 2)

End Sub

'------------------------------------------------------------------------------
'   Make beginning and end of child menu link to page ppage
'   of parent menu
'------------------------------------------------------------------------------
'
Sub MakeChildMenu (child, parent, ppage As Integer)

    Dim x As Integer

    x = menus(child).pagecount

    menus(child).page(0).vectors(0, 2) = parent
    menus(child).page(0).vectors(1, 2) = ppage
    menus(child).page(x - 1).vectors(0, 5) = parent
    menus(child).page(x - 1).vectors(1, 5) = ppage

End Sub

'------------------------------------------------------------------------------
'       Make the up and down arrows cycle through all pages of menu m
'------------------------------------------------------------------------------
'
Sub MakeCircular (m As Integer)

	Dim l, x As Integer

	x = menus(m).pagecount

	For l = 0 To x - 1
		menus(m).page(l).vectors(0, 2) = m
		menus(m).page(l).vectors(1, 2) = (l - 1 + x) Mod x
		menus(m).page(l).vectors(0, 5) = m
		menus(m).page(l).vectors(1, 5) = (l + 1) Mod x
	Next l


End Sub

'------------------------------------------------------------------------------
'       Repair an open fault for channel (index)
'------------------------------------------------------------------------------
'
Sub MakeClosed (index)
	
	Channels(index).errnum = 0
	SensorForm!Open1Box(index).Value = 0

End Sub

'------------------------------------------------------------------------------
'       Simulate an open fault for channel (index)
'------------------------------------------------------------------------------
'
Sub MakeOpen (index)
	
	Channels(index).errnum = 3
	SensorForm!Open1Box(index).Value = 1

End Sub

'------------------------------------------------------------------------------
'       Simulate a short fault for channel (index)
'------------------------------------------------------------------------------
'
Sub makeshort (index)
	
	Channels(index).errnum = 4
	SensorForm!Short1Box(index).Value = 1

End Sub

'------------------------------------------------------------------------------
'       Check gap and faults for cahnnel ch
'------------------------------------------------------------------------------
'
Sub ProcessChannel (ch As Integer)
    Dim gap As Single
    
    gap = SensorForm!Gap1Bar(ch).Value / 100
    Channels(ch).gap = gap
    Channels(ch).errnum = 0
    
    If SensorForm.Open1Box(ch).Value = 1 Then
		Channels(ch).errnum = 3
    End If

    If SensorForm.Short1Box(ch).Value = 1 Then
		Channels(ch).errnum = 4
    End If

    If Channels(ch).errnum < 3 Then  ' no fatal error
		If gap < Channels(ch).mingap Then
		    Channels(ch).errnum = 1
		End If
		If gap < THRESHOLD And gap > Channels(ch).maxgap Then
		    Channels(ch).errnum = 2
		End If
    End If
	
End Sub

'------------------------------------------------------------------------------
' Send a string to the remote display (on COMPORT)
' Uses INT 14 function 1
'------------------------------------------------------------------------------
'
Sub sendchars (chars As String)
    Dim regs As VBregs
    
	If REMOTE Then
	    For i = 1 To Len(chars)
			regs.ax = &H100 + Asc(Mid$(chars, i, 1))
			regs.dx = COMPORT
			Call vbinterruptx(&H14, regs, regs)
	    Next i
	End If

End Sub

'------------------------------------------------------------------------------
'       Send carriage return to remote display
'------------------------------------------------------------------------------
'
Sub sendCR ()

    Dim regs As VBregs

	If REMOTE Then
		regs.ax = &H10D
		regs.dx = COMPORT
		Call vbinterruptx(&H14, regs, regs)
	End If

End Sub

Sub SetContrast (Contrast As Integer)

    Dim regs As VBregs

	If REMOTE Then
	    regs.ax = &H11B
	    regs.dx = COMPORT
	    Call vbinterruptx(&H14, regs, regs)
    
	    regs.ax = &H100 + Asc("L")
	    regs.dx = COMPORT
	    Call vbinterruptx(&H14, regs, regs)
    
	    regs.ax = &H100 + Contrast
	    regs.dx = COMPORT
	    Call vbinterruptx(&H14, regs, regs)
	End If

End Sub

Sub SetParent (m, p, h As Integer)
'------------------------------------------------------------------------------
'       Sets every page of m to return to page p of higher level menu h
'------------------------------------------------------------------------------
	Dim l As Integer

	For l = 0 To menus(m).pagecount - 1
		menus(m).page(l).vectors(0, 6) = p
		menus(m).page(l).vectors(1, 6) = h
	Next l

End Sub

Sub ShowBuffer ()
	
    Dim Text4 As String

    Text4 = ""
    Select Case SystemMenu
	Case TOP_MENU
	    If SystemPage = 2 Then
		Text4 = "Channel " + keybuffer
	    End If
	Case MONITOR_MENU, MONITOR1_MENU, MONITOR2_MENU
	    Text4 = "Channel " + keybuffer
	Case FAULT_BY_LEG_MENU
	    Text4 = "Leg " + keybuffer
	Case FAULT_BY_CHANNEL_MENU
	    Text4 = "CHANNEL " + keybuffer
    End Select

    If Text4 <> "" Then
	Form1.Line4.Text = Text4
    End If
    
End Sub

Sub unshort (index)
	
	Channels(index).errnum = 0
	SensorForm!Short1Box(index).Value = 0

End Sub

Sub UpdatePSmenu (ch As Integer)

    Dim i, m, l As Integer
    Dim gap As Single

    m = PRESENT_STATUS_MENU
    l = ch * 2                             ' menu line number

    gap = Channels(ch).gap

	If Channels(ch).errnum < 3 Then  ' no fatal error
	    If gap < THRESHOLD Then
		menus(m).page(l).line2 = Channels(ch).ntext
	    Else
		menus(m).page(l).line2 = Channels(ch).ftext
	    End If
	Else                            ' fatal error
	    menus(m).page(l).line2 = "State Unknown"
	End If

	l = l + 1                     ' Make next page

	If Channels(ch).errnum = 0 Then  ' no error
	    If gap < THRESHOLD Then
		menus(m).page(l).line2 = "Target Near"
	    Else
		menus(m).page(l).line2 = "Target Far"
	    End If
	Else                            ' error
	    menus(m).page(l).line2 = ErrText(Channels(ch).errnum)
	End If
	
End Sub

Sub UpdateRTmenu (ch As Integer)
    Dim i, m, l As Integer
    Dim gap As Single
    Dim gaptext, Text3, Text4 As String

    If ch < 3 Then
	m = MONITOR1_MENU
	l = ch                              ' menu line number
    Else
	If ch < 6 Then
	    m = MONITOR2_MENU
	    l = ch - 3
	Else
	    m = MONITOR3_MENU
	    l = ch - 6
	End If
    End If

    gap = Channels(ch).gap

    If Channels(ch).errnum > 2 Then  ' fatal error
	Text3 = "State Unknown"
	Text4 = ErrText(Channels(ch).errnum)
    Else
	If Channels(ch).errnum > 0 Then                          ' no fatal error
	    Text3 = ErrText(Channels(ch).errnum)
	Else
	    If gap < THRESHOLD Then
		Text3 = "TARGET NEAR"
	    Else
		Text3 = "TARGET FAR"
	    End If
	End If
	If gap <= .3 Then
	    gaptext = "(" + Format$(gap, fixed) + ") "
	Else
	    gaptext = "(>.3) "
	End If

	Text3 = Text3 + Space$(20 - Len(gaptext) - Len(Text3)) + gaptext
	Text4 = Format$(Channels(ch).mingap, fixed)
	Text4 = Text4 + "    " + Format$(Channels(ch).nomgap, fixed)
	Text4 = Text4 + "    " + Format$(Channels(ch).maxgap, fixed)
    
    End If
    menus(m).page(l).line3 = Text3
    menus(m).page(l).line4 = Text4
End Sub

Sub UpdateSOmenu (p As Integer)

	Dim m As Integer

	m = SYSTEM_OUTPUTS_MENU

	Select Case p
		Case 0                          ' left gear safe: no fatal error, gear down
			Text1 = "FALSE"
			If Channels(LGUL).errnum < 3 And Channels(LGDL).errnum < 3 Then
				If Channels(LGDL).gap < THRESHOLD Then
					If Channels(LGUL).gap > THRESHOLD Then
						Text1 = "TRUE"
					End If
				End If
			End If
			menus(m).page(p).line2 = Text1
		Case 1                          ' left gear unsafe: fatal error, gear not up or down
			Text1 = "FALSE"
			If Channels(LGUL).errnum > 2 Or Channels(LGDL).errnum > 2 Then
				Text1 = "TRUE"
			End If
			If Channels(LGUL).gap > THRESHOLD And Channels(LGDL).gap > THRESHOLD Then
				Text1 = "TRUE"
			End If
			menus(m).page(p).line2 = Text1
		Case 2                          ' right gear safe: no fatal error, gear down
			Text1 = "FALSE"
			If Channels(RGUL).errnum < 3 And Channels(RGDL).errnum < 3 Then
				If Channels(RGDL).gap < THRESHOLD Then
					If Channels(RGUL).gap > THRESHOLD Then
						Text1 = "TRUE"
					End If
				End If
			End If
			menus(m).page(p).line2 = Text1
		Case 3                          ' right gear unsafe: fatal error, gear not up or down
			Text1 = "FALSE"
			If Channels(RGUL).errnum > 2 Or Channels(RGDL).errnum > 2 Then
				Text1 = "TRUE"
			End If
			If Channels(RGUL).gap > THRESHOLD And Channels(RGDL).gap > THRESHOLD Then
				Text1 = "TRUE"
			End If
			menus(m).page(p).line2 = Text1
		Case 4                          ' nose gear safe: no fatal error, gear down
			Text1 = "FALSE"
			If Channels(NGUL).errnum < 3 And Channels(NGDL).errnum < 3 Then
				If Channels(NGDL).gap < THRESHOLD Then
					If Channels(NGUL).gap > THRESHOLD Then
						Text1 = "TRUE"
					End If
				End If
			End If
			menus(m).page(p).line2 = Text1
		Case 5                          ' nose gear unsafe: fatal error, gear not up or down
			Text1 = "FALSE"
			If Channels(NGUL).errnum > 2 Or Channels(NGDL).errnum > 2 Then
				Text1 = "TRUE"
			End If
			If Channels(NGUL).gap > THRESHOLD And Channels(NGDL).gap > THRESHOLD Then
				Text1 = "TRUE"
			End If
			menus(m).page(p).line2 = Text1
		Case 6                          ' doors closed
			Text1 = "CLOSED"
			If Channels(LDOOR).gap > THRESHOLD Or Channels(RDOOR).gap > THRESHOLD Then
				Text1 = "NOT CLOSED"
			End If
			If Channels(LDOOR).errnum > 2 Or Channels(RDOOR).errnum > 2 Then
				Text1 = "NOT CLOSED"
			End If
			menus(m).page(p).line2 = Text1
	End Select

End Sub

