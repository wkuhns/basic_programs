VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form RaceForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Race Game"
   ClientHeight    =   6370
   ClientLeft      =   1310
   ClientTop       =   1870
   ClientWidth     =   9610
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "RACEFORM.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   637
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   961
   Begin MSCommLib.MSComm Com2 
      Left            =   8760
      Top             =   0
      _ExtentX        =   670
      _ExtentY        =   670
      _Version        =   327680
      CommPort        =   2
      DTREnable       =   -1  'True
      ParityReplace   =   32
   End
   Begin MSComDlg.CommonDialog TrackDialog 
      Left            =   9120
      Top             =   0
      _ExtentX        =   564
      _ExtentY        =   564
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
      Height          =   6024
      Left            =   0
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   960
      TabIndex        =   0
      Top             =   360
      Width           =   9624
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

Private Sub FileMenu_Click(index As Integer)
    
    Select Case index
        Case 0              ' new track
            TrackDialog.Filter = trkfilter
            TrackDialog.Action = 1
            trackfile = TrackDialog.filename
            Track_load
        Case 1              ' save track
            track_save
        Case 2              ' save track as...
            TrackDialog.Filter = trkfilter
            TrackDialog.filename = "*.trk"
            TrackDialog.Action = 2
            trackfile = TrackDialog.filename
            track_save
        Case 4              ' Quit
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
        cars(i).active = False
        actives(i) = False
    Next i

    cars(0).active = True
    actives(0) = True
    cars(1).active = True
    actives(1) = True

    trackfile = "track1.trk"
    Track_load

End Sub

Private Sub GameMenu_Click(index As Integer)

    Select Case index
        Case 0                  ' start
            StartGame
        Case 1                  ' show HOF window
            HOFForm.Show
    End Select

End Sub

Private Sub junkbox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Debug.Print "mOUSE DOWN"
End Sub

Private Sub RemoteMenu_Click(index As Integer)

    Select Case index
        Case 0              ' Set Master
            remotestatus = True
            master = True
            OpenCom
        Case 1
            remotestatus = True
            master = False
            OpenCom
            StartGame
        Case 3              ' no remote
            remotestatus = False
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

