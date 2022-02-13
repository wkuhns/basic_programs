VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form RaceForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "  Racetrack"
   ClientHeight    =   6372
   ClientLeft      =   1308
   ClientTop       =   1872
   ClientWidth     =   9612
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "raceform.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   801
   Begin VB.TextBox PlayerBox 
      Height          =   288
      Index           =   3
      Left            =   6480
      TabIndex        =   8
      Top             =   480
      Width           =   1932
   End
   Begin VB.TextBox PlayerBox 
      Height          =   288
      Index           =   2
      Left            =   4440
      TabIndex        =   7
      Top             =   480
      Width           =   1932
   End
   Begin VB.TextBox PlayerBox 
      Height          =   288
      Index           =   1
      Left            =   2400
      TabIndex        =   6
      Top             =   480
      Width           =   1932
   End
   Begin VB.TextBox PlayerBox 
      Height          =   288
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   1932
   End
   Begin VB.TextBox ChatBox 
      Height          =   372
      Left            =   6000
      TabIndex        =   4
      Top             =   0
      Width           =   3612
   End
   Begin MSWinsockLib.Winsock Netsocket 
      Index           =   0
      Left            =   8400
      Top             =   0
      _ExtentX        =   593
      _ExtentY        =   593
   End
   Begin MSComDlg.CommonDialog TrackDialog 
      Left            =   9120
      Top             =   0
      _ExtentX        =   572
      _ExtentY        =   572
      _Version        =   327680
      FontSize        =   0
      MaxFileSize     =   256
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   372
      Left            =   1200
      TabIndex        =   3
      Top             =   0
      Width           =   1212
   End
   Begin VB.TextBox MessageBox 
      Appearance      =   0  'Flat
      Height          =   372
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   3600
   End
   Begin VB.TextBox NameBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   372
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1212
   End
   Begin VB.PictureBox Track 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5544
      Left            =   0
      ScaleHeight     =   460
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   840
      Width           =   9624
   End
   Begin MSWinsockLib.Winsock Netsocket 
      Index           =   1
      Left            =   0
      Top             =   0
      _ExtentX        =   593
      _ExtentY        =   593
   End
   Begin MSWinsockLib.Winsock Netsocket 
      Index           =   2
      Left            =   0
      Top             =   0
      _ExtentX        =   593
      _ExtentY        =   593
   End
   Begin MSWinsockLib.Winsock Netsocket 
      Index           =   3
      Left            =   0
      Top             =   0
      _ExtentX        =   593
      _ExtentY        =   593
   End
   Begin MSWinsockLib.Winsock Netsocket 
      Index           =   4
      Left            =   0
      Top             =   0
      _ExtentX        =   593
      _ExtentY        =   593
   End
   Begin VB.Menu TopMenu 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu FileMenu 
         Caption         =   "&Load Track"
         Index           =   0
      End
      Begin VB.Menu FileMenu 
         Caption         =   "&Save Track"
         Index           =   1
      End
      Begin VB.Menu FileMenu 
         Caption         =   "Save Track &As..."
         Index           =   2
      End
      Begin VB.Menu FileMenu 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu FileMenu 
         Caption         =   "&Quit"
         Index           =   4
      End
   End
   Begin VB.Menu TopMenu 
      Caption         =   "&Game"
      Index           =   1
      Begin VB.Menu GameMenu 
         Caption         =   "Start"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu GameMenu 
         Caption         =   "Hall of Fame"
         Index           =   1
      End
   End
   Begin VB.Menu TopMenu 
      Caption         =   "&Players"
      Index           =   2
   End
   Begin VB.Menu TopMenu 
      Caption         =   "&Design"
      Index           =   3
      Begin VB.Menu TrackMenu 
         Caption         =   "New Picture"
         Index           =   0
      End
      Begin VB.Menu TrackMenu 
         Caption         =   "Setup Track"
         Index           =   1
      End
   End
   Begin VB.Menu TopMenu 
      Caption         =   "&Remote"
      Index           =   4
      Begin VB.Menu RemoteMenu 
         Caption         =   "Set as Master"
         Index           =   0
      End
      Begin VB.Menu RemoteMenu 
         Caption         =   "Set as Slave"
         Index           =   1
      End
      Begin VB.Menu RemoteMenu 
         Caption         =   "No Remote"
         Index           =   2
      End
   End
End
Attribute VB_Name = "RaceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intMax As Long


Private Sub FileMenu_Click(index As Integer)
    
    Dim i As Integer
    
    Select Case index
        Case 0              ' new track
            TrackDialog.Filter = trkfilter
            TrackDialog.Action = 1
            trackfile = TrackDialog.filename
            Track_load
            If Master Then
                netcast "Track: " + trackfile + "|"
            End If
        Case 1              ' save track
            track_save
        Case 2              ' save track as...
            TrackDialog.Filter = trkfilter
            TrackDialog.filename = "*.trk"
            TrackDialog.Action = 2
            trackfile = TrackDialog.filename
            track_save
        Case 4              ' Quit
    
            For i = 0 To intMax
                Netsocket(i).Close
            Next i
    
            End
    End Select

End Sub

Private Sub Form_Load()
    
    Dim i As Integer

    cars(0).color = RGB(255, 0, 0)
    cars(1).color = RGB(0, 255, 0)
    cars(2).color = RGB(0, 0, 255)
    cars(3).color = RGB(255, 255, 0)
    cars(4).color = RGB(255, 0, 255)

    cars(0).name = "Johnathan"
    cars(1).name = "Jack"
    cars(2).name = "Matt"
    cars(3).name = "Sally"
    cars(4).name = "Adam"

    For i = 0 To 4
        cars(i).local = True
        cars(i).status = inactive
    Next i

    trackfile = "track1.trk"
    Track_load

End Sub

Private Sub Form_Terminate()
    Dim i As Integer
    
    MsgBox ("Form Terminating")
    For i = 0 To intMax
        Netsocket(i).Close
    Next i
    
End Sub

Private Sub GameMenu_Click(index As Integer)

    Select Case index
        Case 0                  ' start
            StartGame
        Case 1                  ' show HOF window
            HOFForm.Show
    End Select

End Sub

Sub netcast(msg As String)

    Dim i As Integer
    
    For i = 1 To intMax
        Netsocket(i).SendData (msg)
    Next i
    
End Sub

Private Sub Netsocket_ConnectionRequest(index As Integer, ByVal requestID As Long)
' We've got a request to connect from a client. Give them the next slot
    
    If index = 0 Then
        intMax = intMax + 1
        'Load Netsocket(intMax)
        Netsocket(intMax).LocalPort = 0
        Netsocket(intMax).Accept requestID
    End If
End Sub
    

Private Sub Netsocket_DataArrival(index As Integer, ByVal bytesTotal As Long)
    
    Dim combuff As String
    Dim eom As Integer
    
    Netsocket(index).GetData combuff, vbString, bytesTotal
    
    RaceForm.ChatBox.Text = combuff
    
    ' Parse and decide what to do....

parse:
    eom = InStr(combuff, "|")
    
    ' the '|' character is a divider. Don't pass it on...
    
    ParseMessage Left$(combuff, eom - 1)

    ' if there's another part, parse it...
    
    If eom < (Len(combuff) - 1) Then
        combuff = Right$(combuff, bytesTotal - eom)
        GoTo parse
    End If
    
End Sub

Sub ParseMessage(combuff As String)

    Dim pname As String
    Dim grid As Integer
    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
    
    If InStr(1, combuff, "Player: ") Then
        pname = Mid$(combuff, 9, Len(combuff) - 8)
        'MsgBox ("Enrolling " + pname)
        grid = Val(Mid$(pname, 1, 1))
        pname = Mid$(pname, 3, Len(pname) - 2)
        Enroll pname, grid, False
    End If
    
    If InStr(1, combuff, "Move: ") Then
        pname = Mid$(combuff, 7, Len(combuff) - 6)
        grid = Val(pname)
        'strip off grid
        pname = Mid$(pname, InStr(1, pname, ",") + 1, Len(pname))
        x = Val(pname)
        'strip off grid
        pname = Mid$(pname, InStr(1, pname, ",") + 1, Len(pname))
        y = Val(pname)
        For i = 0 To 4
            If cars(i).grid = grid Then
                cars(i).xmove = x
                cars(i).ymove = y
                cars(i).status = ready
                'ChatBox.Text = "Netmove " + cars(i).name
                processMove i, x, y
                Exit For
            End If
        Next i
    End If
    
    If InStr(1, combuff, "Track: ") Then
        pname = Mid$(combuff, 8, Len(combuff) - 7)
        trackfile = pname
        Track_load
    End If
    
End Sub

Private Sub Netsocket_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    RaceForm!MessageBox.Text = "Network error: " + Description

End Sub

Private Sub RemoteMenu_Click(index As Integer)

    Select Case index
        Case 0              ' Set server
            Master = True
            Slave = False
            Netsocket(0).Close
            Netsocket(0).LocalPort = 2321
            Netsocket(0).Listen
        Case 1              ' Set client
            Master = False
            Slave = True
            Netsocket(0).Close
            Netsocket(0).RemoteHost = "rrvirtnt.cld.com"
            Netsocket(0).RemotePort = 2321
            Netsocket(0).Connect
            StartGame
        Case 3              ' no remote
            Master = False
            Slave = False
    End Select

End Sub

Private Sub SetupTrack()
    ' start setup process. Balance handled in Track_mousedown
    '
    RaceForm!MessageBox.Text = "Click on barrier"
    track_setup = 6
End Sub

Private Sub TopMenu_Click(index As Integer)
    
    Select Case index

        Case 2                  ' player menu
            PlayerForm.Show
    
    End Select

End Sub

Private Sub Track_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Select Case track_setup
        Case 0                  ' playing game
            TrackClick Int(x), Int(y)
'            If cars(player).local Then
'                ProcessMove Int(X), Int(Y)
'            End If
        Case 1                  ' setting start line
            sl_xf = x
            sl_yf = y
            RaceForm!MessageBox.Text = "Finished with track setup"
            track_setup = 0
        Case 2                  ' setting start line
            sl_xs = x
            sl_ys = y
            RaceForm!MessageBox.Text = "Click on other end of start/finish line"
            track_setup = 1
        Case 3                  ' setting ice
            icecolor = RaceForm!Track.Point(x, y)
            RaceForm!MessageBox.Text = "Click on one end of start/finish line"
            track_setup = 2
        Case 4                  ' setting wet
            wetcolor = RaceForm!Track.Point(x, y)
            RaceForm!MessageBox.Text = "Click on icy part of track"
            track_setup = 3
        Case 5                  ' setting road
            roadcolor = RaceForm!Track.Point(x, y)
            RaceForm!MessageBox.Text = "Click on wet part of track"
            track_setup = 4
        Case 6                  ' setting wet
            wallcolor = RaceForm!Track.Point(x, y)
            RaceForm!MessageBox.Text = "Click on normal part of track"
            track_setup = 5
    End Select
End Sub

Private Sub TrackMenu_Click(index As Integer)
    
    Select Case index
        Case 0                      ' new bitmap
            TrackDialog.Filter = bmpfilter
            TrackDialog.Action = 1
            bmpfile = TrackDialog.filename
            Track.Picture = LoadPicture(bmpfile)
            hofcount = 0
        Case 1                      ' setup
            SetupTrack
    End Select

End Sub


