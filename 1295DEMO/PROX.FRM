VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Remote Unit"
   ClientHeight    =   5508
   ClientLeft      =   96
   ClientTop       =   420
   ClientWidth     =   3288
   Height          =   5928
   Icon            =   PROX.FRX:0000
   Left            =   48
   LinkTopic       =   "Form1"
   Picture         =   PROX.FRX:0302
   ScaleHeight     =   459
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   274
   Top             =   48
   Width           =   3384
   Begin Timer Timer3 
      Interval        =   90
      Left            =   0
      Top             =   5160
   End
   Begin Timer Timer2 
      Interval        =   100
      Left            =   0
      Top             =   4560
   End
   Begin TextBox Line1 
      BorderStyle     =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Courier New"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   372
      Left            =   120
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   120
      Width           =   3012
   End
   Begin TextBox Line2 
      BorderStyle     =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Courier New"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   372
      Left            =   120
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   480
      Width           =   3012
   End
   Begin TextBox Line4 
      BorderStyle     =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Courier New"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   372
      Left            =   120
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3012
   End
   Begin TextBox Line3 
      BorderStyle     =   0  'None
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Courier New"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   372
      Left            =   120
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   840
      Width           =   3012
   End
   Begin CommandButton FBCBtn 
      Caption         =   "FBC"
      Height          =   492
      Left            =   240
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4800
      Width           =   612
   End
   Begin CommandButton FBLBtn 
      Caption         =   "FBL"
      Height          =   492
      Left            =   240
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4200
      Width           =   612
   End
   Begin CommandButton DataBtn 
      Caption         =   "Inputs"
      Height          =   492
      Left            =   240
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3600
      Width           =   612
   End
   Begin CommandButton Btn0 
      Caption         =   "0"
      Height          =   492
      Left            =   960
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4800
      Width           =   612
   End
   Begin CommandButton EnterBtn 
      Caption         =   "Enter"
      Height          =   492
      Left            =   2400
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4800
      Width           =   612
   End
   Begin CommandButton BSBtn 
      Caption         =   "B/S"
      Height          =   492
      Left            =   1680
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4800
      Width           =   612
   End
   Begin CommandButton Btn3 
      Caption         =   "3"
      Height          =   492
      Left            =   2400
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4200
      Width           =   612
   End
   Begin CommandButton Btn2 
      Caption         =   "2"
      Height          =   492
      Left            =   1680
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4200
      Width           =   612
   End
   Begin CommandButton Btn1 
      Caption         =   "1"
      Height          =   492
      Left            =   960
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4200
      Width           =   612
   End
   Begin CommandButton Btn6 
      Caption         =   "6"
      Height          =   492
      Left            =   2400
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3600
      Width           =   612
   End
   Begin CommandButton Btn5 
      Caption         =   "5"
      Height          =   492
      Left            =   1680
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3600
      Width           =   612
   End
   Begin CommandButton Btn4 
      Caption         =   "4"
      Height          =   492
      Left            =   960
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3600
      Width           =   612
   End
   Begin CommandButton Btn9 
      Caption         =   "9"
      Height          =   492
      Left            =   2400
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3000
      Width           =   612
   End
   Begin CommandButton Btn8 
      Caption         =   "8"
      Height          =   492
      Left            =   1680
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3000
      Width           =   612
   End
   Begin CommandButton Btn7 
      Caption         =   "7"
      Height          =   492
      Left            =   960
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3000
      Width           =   612
   End
   Begin CommandButton MonBtn 
      Caption         =   "RigHelp"
      Height          =   492
      Left            =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3000
      Width           =   612
   End
   Begin CommandButton NoButton 
      Caption         =   "No"
      Height          =   492
      Left            =   960
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2400
      Width           =   612
   End
   Begin CommandButton DnCtrstBtn 
      Caption         =   "-C"
      Height          =   492
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2400
      Width           =   612
   End
   Begin CommandButton UpCtrstButton 
      Caption         =   "+C"
      Height          =   492
      Left            =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
      Width           =   612
   End
   Begin CommonDialog CMDialog1 
      DefaultExt      =   ".dat"
      DialogTitle     =   "Select data file"
      Filter          =   "Data |*.dat|all|*.*"
      InitDir         =   "c:\prox"
      Left            =   0
      Top             =   3360
   End
   Begin Timer Timer1 
      Interval        =   20
      Left            =   0
      Top             =   3960
   End
   Begin CommandButton MenuButton 
      Caption         =   "Menu"
      Height          =   492
      Left            =   2400
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2400
      Width           =   612
   End
   Begin CommandButton DownButton 
      Caption         =   "Down"
      Height          =   492
      Left            =   1680
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2400
      Width           =   612
   End
   Begin CommandButton OnButton 
      Caption         =   "On"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   7.8
      FontStrikethru  =   0   'False
      FontUnderline   =   -1  'True
      Height          =   492
      Left            =   2400
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
      Width           =   612
   End
   Begin CommandButton UpButton 
      Caption         =   "Up"
      Height          =   492
      Left            =   1680
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1800
      Width           =   612
   End
   Begin CommandButton YesButton 
      Caption         =   "Yes"
      Height          =   492
      Left            =   960
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1800
      Width           =   612
   End
   Begin Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   1212
      Left            =   900
      Top             =   1740
      Width           =   2172
   End
End

Declare Sub vbinterruptx Lib "VBASM.DLL" (ByVal IntNum As Integer, InRegs As VBregs, OutRegs As VBregs)

Sub BSBtn_Click ()

    If Len(keybuffer) > 0 Then
	keybuffer = Left$(keybuffer, Len(keybuffer) - 1)
    End If
    ShowBuffer

End Sub

Sub Btn0_Click ()

    keybuffer = keybuffer + "0"
    ShowBuffer

End Sub

Sub Btn1_Click ()

    keybuffer = keybuffer + "1"
    ShowBuffer

End Sub

Sub Btn2_Click ()

    keybuffer = keybuffer + "2"
    ShowBuffer

End Sub

Sub Btn3_Click ()

    keybuffer = keybuffer + "3"
    ShowBuffer

End Sub

Sub Btn4_Click ()

    keybuffer = keybuffer + "4"
    ShowBuffer

End Sub

Sub Btn5_Click ()

    keybuffer = keybuffer + "5"
    ShowBuffer

End Sub

Sub Btn6_Click ()

    keybuffer = keybuffer + "6"
    ShowBuffer

End Sub

Sub Btn7_Click ()

    keybuffer = keybuffer + "7"
    ShowBuffer

End Sub

Sub Btn8_Click ()

    keybuffer = keybuffer + "8"
    ShowBuffer

End Sub

Sub Btn9_Click ()

    keybuffer = keybuffer + "9"
    ShowBuffer

End Sub

Sub DnCtrstBtn_Click ()
    
    Dim contrast As Integer

    contrast = form3.ContrastBar.Value - 5

    If contrast < 64 Then
	contrast = 64
    End If

    form3.ContrastBar = contrast

    SetContrast (contrast)

End Sub

Sub DownButton_Click ()
    ProcessVector (5)
    keybuffer = ""
End Sub

Sub EnterBtn_Click ()
    
    Dim choice As Integer

    choice = Val(keybuffer)

    Select Case SystemMenu
	Case TOP_MENU
	    If SystemPage = 5 Then
		If choice < 4 Then
		    SystemMenu = MONITOR1_MENU
		    SystemPage = choice - 1
		Else
		    If choice < 7 Then
			SystemMenu = MONITOR2_MENU
			SystemPage = choice - 4
		    Else
			SystemMenu = MONITOR3_MENU
			SystemPage = choice - 7
		    End If
		End If
	    End If
	Case MONITOR_MENU, MONITOR1_MENU, MONITOR2_MENU, MONITOR3_MENU
		If choice < 4 Then
		    SystemMenu = MONITOR1_MENU
		    SystemPage = choice - 1
		Else
		    If choice < 7 Then
			SystemMenu = MONITOR2_MENU
			SystemPage = choice - 4
		    Else
			SystemMenu = MONITOR3_MENU
			SystemPage = choice - 7
		    End If
		End If
    End Select
    keybuffer = ""

End Sub

Sub FBCBtn_Click ()
    SystemMenu = FAULT_BY_CHANNEL_MENU
    SystemPage = 0
    keybuffer = ""
End Sub

Sub FBLBtn_Click ()

    SystemMenu = FAULT_BY_LEG_MENU
    SystemPage = 0
    keybuffer = ""
    
End Sub

Sub Form_Load ()

    Dim i, count As Integer
    
    InitPort                        ' Set up BIOS serial port parameters
    
    ' Zero out all key vectors
    
    For i = 1 To MAXMENU
       Menus(i).pagecount = 0
       For j = 0 To MAXMSG - 1
	   For k = 1 To 6
		Menus(i).page(j).vectors(0, k) = 0
		Menus(i).page(j).vectors(1, k) = 0
	   Next k
       Next j
    Next i
    
    On Error Resume Next
    Open "prox.cfg" For Input As #1
    
    If Err > 0 Then
	On Error GoTo 0
	Open "n:\demo\prox.cfg" For Input As #1
    End If
    On Error GoTo 0

    Input #1, count
    For i = 1 To count
	Input #1, ErrText(i)
    Next i

    Input #1, count
    For i = 0 To count - 1
	Input #1, Channels(i).name, Channels(i).abbr, Channels(i).ntext
	Input #1, Channels(i).ftext, Channels(i).mingap, Channels(i).nomgap
	Input #1, Channels(i).maxgap
    Next i

    Close #1

    InitPSmenu
    InitSOmenu
    InitRTmenu
    ClearRemote

    BacklightOff
    
    For i = 0 To NUMCHANNELS - 1
	GapChange (i)
    Next i

    form3.Show
    sensorform.Show

End Sub

Sub MenuButton_Click ()
    ProcessVector (6)
    keybuffer = ""
End Sub

Sub MonBtn_Click ()

    SystemMenu = 18
    SystemPage = 0
    keybuffer = ""

End Sub

Sub NoButton_Click ()

    ProcessVector (4)
    keybuffer = ""
End Sub

' Starts system. The datafile is read, the display is
' turned on, and the diagnostic window is displayed.
' If the system is alreday on, it's turned off. (display
' is blanked, really.)
'
Sub OnButton_Click ()

  Dim keyname As String
  Dim m, c As Integer

  If SystemMenu = 0 Then
    OnButton.Caption = "Off"
    
    ' Use Common Dialog to select file to read for
    ' menu structure.
    ' default is c:\prox\*.dat

    '    CMDialog1.Action = 1
    '    Open CMDialog1.Filename For Input As #1

    On Error Resume Next
    Open "qterm2.dat" For Input As #1
    
    If Err > 0 Then
	On Error GoTo 0
	Open "n:\demo\1295demo\qterm2.dat" For Input As #1
    End If
    On Error GoTo 0

    While Not EOF(1)
	Input #1, menu, choice
	choice = choice - 1
	Input #1, Menus(menu).page(choice).line1
	Input #1, Menus(menu).page(choice).line2
	Input #1, Menus(menu).page(choice).line3
	Input #1, Menus(menu).page(choice).line4
	Input #1, keyname
	While Len(keyname) > 0
	    Input #1, m, c
	    c = c - 1
	    Select Case keyname
	    Case "YES"
		Menus(menu).page(choice).vectors(0, 1) = m
		Menus(menu).page(choice).vectors(1, 1) = c
	    Case "UP"
		Menus(menu).page(choice).vectors(0, 2) = m
		Menus(menu).page(choice).vectors(1, 2) = c
	    Case "ON"
		Menus(menu).page(choice).vectors(0, 3) = m
		Menus(menu).page(choice).vectors(1, 3) = c
	    Case "NO"
		Menus(menu).page(choice).vectors(0, 4) = m
		Menus(menu).page(choice).vectors(1, 4) = c
	    Case "DOWN"
		Menus(menu).page(choice).vectors(0, 5) = m
		Menus(menu).page(choice).vectors(1, 5) = c
	    Case "MENU"
		Menus(menu).page(choice).vectors(0, 6) = m
		Menus(menu).page(choice).vectors(1, 6) = c
	    End Select
	    Input #1, keyname
	Wend
    Wend

    Close #1
    
    ' Set global variables to menu 1, line 1
    ' process vector 6, which is 'menu' key

    SystemMenu = TOP_MENU
    SystemPage = 0
    
    BacklightOn
    keybuffer = ""

    line1.Text = Menus(1).page(0).line1$
    line2.Text = Menus(1).page(0).line2$
    line3.Text = Menus(1).page(0).line3$
    line4.Text = Menus(1).page(0).line4$
    Dform.Show
    ProcessVector (6)   ' act like 'menu' was pressed
  Else                  ' already 'on', so turn 'off'
    SystemMenu = 0
    line1.Text = ""
    line2.Text = ""
    line3.Text = ""
    line4.Text = ""
    OnButton.Caption = "On"
	
    ' clear remote display with ESC H ESC J
    
    ClearRemote

    BacklightOff
    
    Dform.Hide          ' make debug window disappear
    keybuffer = ""
  
  End If

End Sub

' A button was pressed. Figure out what, if anything, we
' should do about it. The global variables SystemMenu
' and SystemPage tell us the current menu and line. We
' look in the global Menus structure for the vector entry
' corresponding to the key that was pressed.
'
Sub ProcessVector (button)
    
    ' extract destination menu (m) and line (l)

    m = Menus(SystemMenu).page(SystemPage).vectors(0, button)
    l = Menus(SystemMenu).page(SystemPage).vectors(1, button)
    If m <> 0 Then      ' nonzero means do something
	SystemMenu = m
	SystemPage = l
    End If

End Sub

' This function is invoked by the timer at 50 hz.
' The serial port is checked for available keystroke(s)
' If there is a character, it is read and processed.
'
Sub Timer1_Timer ()

    Dim regs As VBregs
    
	If REMOTE Then
	    ' Int 14 function 3: check serial port status
	    ' Bit 0 of AH is data available flag

	    regs.ax = &H300
	    regs.dx = COMPORT
	    Call vbinterruptx(&H14, regs, regs)
	    If (regs.ax And &H100) Then
	
		' There's a keystroke - let's read it.
		' Int 14 function 2 reads character

		regs.ax = &H200
		regs.dx = COMPORT
		Call vbinterruptx(&H14, regs, regs)
		keypress = regs.ax
	
		' Hairbrain key mapping is from random keypad
		' connections on prototype.
		' Transmitted keycodes are set in the qterm
		' data file \qterm\qdataiii.dat which is loaded
		' into the qterm with the qsetup.exe program.
		' For the remote QTERMII, the data file is
		' \qterm\remote.dat

		Select Case keypress
			Case Asc("d")                   'on/off
			    OnButton_Click
			Case Asc("b")                   'yes
			    YesButton_Click
			Case Asc("g")                   'down
			    DownButton_Click
			Case Asc("c")                   'up
			    UpButton_Click
			Case Asc("h")                   'menu
			    MenuButton_Click
			Case Asc("f")                   'no
			    NoButton_Click
			Case BS                         'backspace
			    BSBtn_Click
			Case Asc("a")                   'contrast up
			    UpCtrstButton_Click
			Case Asc("e")                   'contrast down
			    DnCtrstBtn_Click
			Case CR                         'Enter
			    EnterBtn_Click
			Case Asc("i")                   'RealTime Monitor
			    MonBtn_Click
			Case Asc("k")                   'FaultByLeg
			    FBLBtn_Click
			Case Asc("l")                   'FaultByChannel
			    FBCBtn_Click
			End Select
	    End If
	End If
End Sub

Sub Timer2_Timer ()
    DisplayPage
End Sub

Sub Timer3_Timer ()
    
    Select Case SystemMenu
	Case SYSTEM_OUTPUTS_MENU
	    ProcessChannel (Int(SystemPage / 2))
	    ProcessChannel (Int(SystemPage / 2) + 1)
	    UpdateSOmenu (SystemPage)
	Case PRESENT_STATUS_MENU
	    ProcessChannel (Int(SystemPage / 2))
	    UpdatePSmenu (Int(SystemPage / 2))
	Case MONITOR1_MENU
	    ProcessChannel (SystemPage)
	    UpdateRTmenu (SystemPage)
	Case MONITOR2_MENU
	    ProcessChannel (SystemPage + 3)
	    UpdateRTmenu (SystemPage + 3)
	Case MONITOR3_MENU
	    ProcessChannel (SystemPage + 6)
	    UpdateRTmenu (SystemPage + 6)
    End Select

End Sub

Sub UpButton_Click ()
    ProcessVector (2)
    keybuffer = ""
End Sub

Sub UpCtrstButton_Click ()

    Dim contrast As Integer

    contrast = form3.ContrastBar.Value + 5

    If contrast > 127 Then
	contrast = 127
    End If

    form3.ContrastBar = contrast

    SetContrast (contrast)

End Sub

Sub YesButton_Click ()
    
    ProcessVector (1)
    keybuffer = ""

End Sub

