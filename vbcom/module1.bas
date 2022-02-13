
Sub Delay (amount As Single)
    
    t! = Timer
    
    While t! + amount > Timer
    Wend

End Sub

Sub UpdateCaption (Msg$, Wait As Single)

    Dim wHeight As Integer
    Dim wCenter As Integer

    If CommDemo.TextWidth(CaptionText$) > CommDemo.TextWidth(Msg$) Then

        CommDemo.CurrentX = CaptionLeft
        CommDemo.CurrentY = CaptionCenter
        CommDemo.ForeColor = CommDemo.BackColor
        CommDemo.Print CaptionText$;
        CommDemo.ForeColor = 0

    End If
    
    wHeight = CommDemo.TextHeight(Msg$)
    wCenter = (CaptionHeight - wHeight) / 2

    CaptionCenter = CaptionTop + wCenter
    CaptionText$ = Msg$
    
    CommDemo.CurrentX = CaptionLeft
    CommDemo.CurrentY = CaptionCenter
    CommDemo.Print CaptionText$;

    If Wait Then
        Delay Wait
    End If

End Sub

Function ReadCommPort (ReadAmount As Integer) As String
    
    Dim ApiErr As Integer
    Dim EventMask As Integer
    Dim Found As Integer

    If ReadAmount < 1 Then
        ReadCommPort = ""
        Exit Function
    End If

    EventMask = CommEventMask
    ApiErr = GetCommEventMask(CommHandle, EventMask)
    
    If ApiErr And EV_RXCHAR Then
        Buffer$ = Space$(ReadAmount)
        ApiErr = ReadComm(CommHandle, Buffer$, Len(Buffer$))

        If ApiErr < 0 Then
            UpdateCaption " ReadCOMM API FAILED! (ERR " + Str$(ApiErr) + ")", 3
            Buffer$ = ""
        Else
            Buffer$ = Left$(Buffer$, ApiErr)
            
            ' Expand CR to CR/LF for "Text" box display

            Found = 1
            Do
                Found = InStr(Found, Buffer$, Chr$(13))
                If Found Then
                    Buffer$ = Left$(Buffer$, Found) + Chr$(10) + Right$(Buffer$, Len(Buffer$) - Found)
                    Found = Found + 1
                End If
            Loop While Found
        End If
    End If

    If (ApiErr And EV_RXFLAG) And (CommEventMask And EV_RXFLAG) Then
    End If

    If (ApiErr And EV_TXEMPTY) And (CommEventMask And EV_XFLAG) Then
    End If

    If (ApiErr And EV_CTS) And (CommEventMask And EV_CTS) Then
    End If

    If (ApiErr And EV_DSR) And (CommEventMask And EV_DSR) Then
    End If

    If (ApiErr And EV_RLSD) And (CommEventMask And EV_RLSD) Then
    End If

    If (ApiErr And EV_BREAK) And (CommEventMask And EV_BREAK) Then
    End If

    If (ApiErr And EV_ERR) And (CommEventMask And EV_ERR) Then
    End If
    
    If (ApiErr And EV_PERR) And (CommEventMask And EV_PERR) Then
    End If
    
    If (ApiErr And EV_RING) And (CommEventMask And EV_RING) Then
        UpdateCaption " Receive Window: RING! ", 0
        Beep
    End If
    
    ReadCommPort = Buffer$

End Function

Sub WriteCommPort (Send$)

    ApiErr% = WriteComm(CommHandle, Send$, Len(Send$))

    If ApiErr% < 0 Then
        UpdateCaption " WriteComm API Failed! (ERR " + Str$(ApiErr%) + ")", 2
    End If

End Sub

Sub DisplayQBOpen (TempDCB As CommStateDCB, DevName As String, RB As Integer, TB As Integer, Interval As Integer)

    ParityChar$ = "NOEMS"

    A$ = " Open " + Chr$(34) + DevName
    A$ = A$ + LTrim$(Str$(TempDCB.BaudRate)) + ","
    A$ = A$ + Mid$(ParityChar$, Asc(TempDCB.Parity) + 1, 1) + ","
    A$ = A$ + LTrim$(Str$(Asc(TempDCB.ByteSize))) + ","
    
    Select Case Asc(TempDCB.StopBits)
        Case 0
            B$ = "1"
        Case 1
            B$ = "1.5"
        Case 2
            B$ = "2"
        Case Else
    End Select

    A$ = A$ + B$ + ","
    
    A$ = A$ + "RB" + LTrim$(Str$(RB)) + ","
    A$ = A$ + "TB" + LTrim$(Str$(TB)) + ","
    A$ = A$ + "CD" + LTrim$(Str$(TempDCB.RlsTimeOut)) + ","
    A$ = A$ + "CS" + LTrim$(Str$(TempDCB.CtsTimeOut)) + ","
    A$ = A$ + "DS" + LTrim$(Str$(TempDCB.DsrTimeOut)) + ","
    A$ = A$ + "TI" + LTrim$(Str$(Interval))
    
    A$ = A$ + Chr$(34)

    UpdateCaption A$, 0

End Sub

Sub Remove_Items_From_SysMenu (A_Form As Form)

    HSysMenu = GetSystemMenu(A_Form.Hwnd, 0)
  
    R = RemoveMenu(HSysMenu, 8, MF_BYPOSITION) 'Switch to
    R = RemoveMenu(HSysMenu, 7, MF_BYPOSITION) 'Separator
    R = RemoveMenu(HSysMenu, 5, MF_BYPOSITION) 'Separator
    R = RemoveMenu(HSysMenu, 4, MF_BYPOSITION) 'Maximize
    R = RemoveMenu(HSysMenu, 3, MF_BYPOSITION) 'Minimize
    R = RemoveMenu(HSysMenu, 2, MF_BYPOSITION) 'Size
    R = RemoveMenu(HSysMenu, 0, MF_BYPOSITION) 'Restore

End Sub

Sub CenterDialog (A_Form As Form)

    Dim cLeft As Integer
    Dim cTop As Integer

    cLeft = (Screen.Width - A_Form.Width) / 2
    cTop = (Screen.Height - A_Form.Height) / 2

    A_Form.Move cLeft, cTop

End Sub

Sub Draw3d (wLeft As Integer, wTop As Integer, wWidth As Integer, wHeight As Integer, A_Form As Form)
    Dim LeftY As Integer
    Dim LeftX As Integer
    
    Dim RightY As Integer
    Dim RightX As Integer

    Dim Depth As Integer

    Dim OffSet As Integer
    Dim SetIn As Integer

    OffSet = 15
    SetIn = 1
    
    ' Draw the Black and White lines to give a "Set In" effect
    ' around the text and buttons

    For Depth = OffSet To OffSet * SetIn Step OffSet
        
        LeftX = wLeft - Depth
        LeftY = wTop - Depth
        RightX = wLeft + wWidth + Depth
        RightY = wTop + wHeight + Depth

        ' Draw the Top and Bottom Lines
        A_Form.Line (LeftX, LeftY)-(RightX, LeftY), QBColor(0)
        A_Form.Line (LeftX, RightY)-(RightX, RightY), QBColor(15)
        
        ' Draw the Left and Right Lines
        A_Form.Line (LeftX - OffSet, LeftY)-(LeftX - OffSet, RightY + OffSet), QBColor(0)
        A_Form.Line (RightX, LeftY)-(RightX, RightY + OffSet), QBColor(15)

    Next Depth

End Sub

