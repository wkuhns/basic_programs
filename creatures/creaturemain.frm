VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form MainForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   Caption         =   "WorldView"
   ClientHeight    =   6036
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   10860
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   6036
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   420
      TabIndex        =   40
      Top             =   1020
      Width           =   4872
      Begin VB.OptionButton SpeedBtn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   192
         Index           =   0
         Left            =   1380
         TabIndex        =   45
         Top             =   60
         Value           =   -1  'True
         Width           =   192
      End
      Begin VB.OptionButton SpeedBtn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   192
         Index           =   1
         Left            =   2100
         TabIndex        =   44
         Top             =   60
         Width           =   192
      End
      Begin VB.OptionButton SpeedBtn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   192
         Index           =   2
         Left            =   2820
         TabIndex        =   43
         Top             =   60
         Width           =   192
      End
      Begin VB.OptionButton SpeedBtn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   192
         Index           =   3
         Left            =   3540
         TabIndex        =   42
         Top             =   60
         Width           =   192
      End
      Begin VB.OptionButton SpeedBtn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Option1"
         Height          =   192
         Index           =   4
         Left            =   4260
         TabIndex        =   41
         Top             =   60
         Width           =   192
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1"
         Height          =   192
         Index           =   0
         Left            =   1620
         TabIndex        =   51
         Top             =   60
         Width           =   312
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "10"
         Height          =   192
         Index           =   1
         Left            =   2340
         TabIndex        =   50
         Top             =   60
         Width           =   312
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "50"
         Height          =   192
         Index           =   2
         Left            =   3060
         TabIndex        =   49
         Top             =   60
         Width           =   312
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "100"
         Height          =   192
         Index           =   3
         Left            =   3780
         TabIndex        =   48
         Top             =   60
         Width           =   312
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "500"
         Height          =   192
         Index           =   4
         Left            =   4560
         TabIndex        =   47
         Top             =   60
         Width           =   312
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Display every"
         Height          =   192
         Index           =   5
         Left            =   0
         TabIndex        =   46
         Top             =   60
         Width           =   1212
      End
   End
   Begin ComctlLib.Slider ClrSlider 
      Height          =   192
      Index           =   0
      Left            =   600
      TabIndex        =   36
      Top             =   5220
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   339
      _Version        =   327680
      Max             =   100
      TickStyle       =   3
   End
   Begin VB.TextBox ClrBox 
      Height          =   264
      Index           =   1
      Left            =   2400
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   5460
      Width           =   492
   End
   Begin VB.TextBox ClrBox 
      Height          =   288
      Index           =   0
      Left            =   600
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   5460
      Width           =   492
   End
   Begin VB.TextBox GFBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   288
      Left            =   6960
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   4560
      Width           =   732
   End
   Begin VB.CommandButton PhysicsBtn 
      Caption         =   "Physics Form"
      Height          =   372
      Left            =   8760
      TabIndex        =   32
      Top             =   480
      Width           =   2052
   End
   Begin VB.PictureBox GeneBox 
      BackColor       =   &H00000000&
      Height          =   1332
      Left            =   7800
      ScaleHeight     =   -100
      ScaleMode       =   0  'User
      ScaleTop        =   100
      ScaleWidth      =   200
      TabIndex        =   29
      Top             =   3480
      Width           =   2892
   End
   Begin VB.PictureBox ChartBox 
      BackColor       =   &H00000000&
      Height          =   1332
      Left            =   7800
      ScaleHeight     =   -100
      ScaleMode       =   0  'User
      ScaleTop        =   100
      ScaleWidth      =   200
      TabIndex        =   27
      Top             =   1440
      Width           =   2892
   End
   Begin VB.TextBox ExhaustionBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   288
      Left            =   6840
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   2520
      Width           =   852
   End
   Begin VB.TextBox BAgeBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   288
      Left            =   6960
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   3840
      Width           =   732
   End
   Begin VB.TextBox BSizeBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   288
      Left            =   6960
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   4200
      Width           =   732
   End
   Begin VB.TextBox AgeBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   288
      Left            =   6840
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   1800
      Width           =   852
   End
   Begin VB.TextBox StarveBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   288
      Left            =   6840
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   2160
      Width           =   852
   End
   Begin VB.TextBox MetBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   288
      Left            =   6960
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3480
      Width           =   732
   End
   Begin VB.TextBox LiveBox 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF80&
      Height          =   288
      Left            =   6840
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   1440
      Width           =   852
   End
   Begin VB.OptionButton MapModeBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Option1"
      Height          =   192
      Index           =   1
      Left            =   4680
      TabIndex        =   11
      Top             =   2220
      Width           =   192
   End
   Begin VB.OptionButton MapModeBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Option1"
      Height          =   192
      Index           =   0
      Left            =   4680
      TabIndex        =   10
      Top             =   1980
      Value           =   -1  'True
      Width           =   192
   End
   Begin VB.TextBox ScaleLo 
      Height          =   288
      Left            =   4680
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   4860
      Width           =   612
   End
   Begin VB.TextBox ScaleHi 
      Height          =   288
      Left            =   4680
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   1380
      Width           =   612
   End
   Begin VB.PictureBox ScaleBox 
      Height          =   3732
      Left            =   4320
      ScaleHeight     =   -3684
      ScaleMode       =   0  'User
      ScaleTop        =   3684
      ScaleWidth      =   204
      TabIndex        =   7
      Top             =   1380
      Width           =   252
   End
   Begin VB.PictureBox MapBox 
      Height          =   3732
      Left            =   360
      ScaleHeight     =   -3684
      ScaleMode       =   0  'User
      ScaleTop        =   3684
      ScaleWidth      =   3684
      TabIndex        =   6
      Top             =   1380
      Width           =   3732
   End
   Begin VB.CommandButton RandBtn 
      Caption         =   "Randomize"
      Height          =   372
      Left            =   8760
      TabIndex        =   5
      Top             =   60
      Width           =   972
   End
   Begin VB.TextBox ClockBox 
      Height          =   252
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   612
   End
   Begin VB.CommandButton RestartBtn 
      Caption         =   "Start Over"
      Height          =   372
      Left            =   9840
      TabIndex        =   3
      Top             =   60
      Width           =   972
   End
   Begin VB.CommandButton StopBtn 
      Caption         =   "Pause"
      Height          =   372
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   972
   End
   Begin VB.CommandButton StartBtn 
      Caption         =   "Run"
      Height          =   372
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   972
   End
   Begin VB.TextBox MessageBox 
      Height          =   852
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   5040
      Width           =   5052
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5040
      Top             =   5520
   End
   Begin ComctlLib.Slider ClrSlider 
      Height          =   192
      Index           =   1
      Left            =   2400
      TabIndex        =   37
      Top             =   5220
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   339
      _Version        =   327680
      Max             =   100
      TickStyle       =   3
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Time (ticks)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   52
      Top             =   120
      Width           =   1092
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   192
      Left            =   3900
      Top             =   5220
      Width           =   192
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   192
      Left            =   2040
      Top             =   5220
      Width           =   192
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   192
      Left            =   360
      Top             =   5220
      Width           =   192
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Graze Factor"
      Height          =   252
      Left            =   5880
      TabIndex        =   35
      Top             =   4560
      Width           =   1092
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Genesis Creature Evolution Project"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   2580
      TabIndex        =   33
      Top             =   240
      Width           =   5952
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Genetic Characteristics"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   6840
      TabIndex        =   31
      Top             =   3120
      Width           =   3732
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Population and causes of death"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   6840
      TabIndex        =   30
      Top             =   1080
      Width           =   3972
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Population"
      Height          =   252
      Left            =   5880
      TabIndex        =   28
      Top             =   1440
      Width           =   852
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Exhaustion"
      Height          =   252
      Left            =   5880
      TabIndex        =   26
      Top             =   2520
      Width           =   852
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Eat Factor"
      Height          =   252
      Left            =   6000
      TabIndex        =   24
      Top             =   3840
      Width           =   852
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Size To Breed"
      Height          =   252
      Left            =   5880
      TabIndex        =   22
      Top             =   4200
      Width           =   1092
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      Height          =   252
      Left            =   5880
      TabIndex        =   20
      Top             =   1800
      Width           =   852
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starvation"
      Height          =   252
      Left            =   5880
      TabIndex        =   18
      Top             =   2160
      Width           =   852
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Metabolism"
      Height          =   252
      Left            =   6000
      TabIndex        =   16
      Top             =   3480
      Width           =   852
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fertility"
      Height          =   252
      Left            =   4920
      TabIndex        =   13
      Top             =   2220
      Width           =   612
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Food"
      Height          =   252
      Left            =   4920
      TabIndex        =   12
      Top             =   1980
      Width           =   612
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   1932
      Left            =   5760
      Top             =   3000
      Width           =   5052
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   1932
      Left            =   5760
      Top             =   960
      Width           =   5052
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   4932
      Left            =   240
      Top             =   960
      Width           =   5292
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ChartData(7, 200) As Single
Dim chartpoint As Integer

Dim MapMode As Integer
Dim RunFast As Boolean
Dim AgeCounter As Long
Dim StarveCounter As Long
Dim ExhaustionCounter As Long
Dim AgeDeaths(100) As Integer
Dim StarveDeaths(100) As Integer
Dim ExhaustionDeaths(100) As Integer
Dim deathcounter As Integer
Dim Every As Integer

Private Sub ClrSlider_Change(Index As Integer)

    ClrBox(Index) = ClrSlider(Index).Value
    
End Sub
' First time form is opened. Do first init here
'
Private Sub Form_Initialize()

    WorldSize = 40
    LifeSize = 150
    
End Sub

' Start the world - I want to get on
'
Private Sub Form_Load()

    MakePhysics         ' Set physical constants
    MakeWorld           ' Set terrain constants and variables
    MakeCreatures       ' Populate the world
    Ticks = 0           ' time = 0
    GuiInit
    
End Sub
Private Sub GuiInit()
    ' Set some GUI stuff
    ' Scale bar limits
    ScaleHi = 25
    ScaleLo = 0
    DrawMapScale        ' Draw vertical scale bar for map
    LiveBox = LiveCount
    AgeBox = 0
    StarveBox = 0
    ExhaustionBox = 0
    ClrSlider(0).Value = 20
    ClrSlider(1).Value = 75
    Every = 1           ' display every tick
    
End Sub
' Draw (or redraw) map scale bar
'
Private Sub DrawMapScale()

    Dim i As Integer
    
    For i = 0 To 10
        ScaleBox.AutoRedraw = True
        
        Select Case MapMode
            Case 0
                ' Print green scale for food
                ScaleBox.Line (0, i * ScaleBox.Height / 11)-Step(ScaleBox.Height / 11, ScaleBox.Height / 11), RGB(0, shade(CSng(i), 0, 11), 0), BF
            Case 1
                ' Print blue scale for fertility
                ScaleBox.Line (0, i * ScaleBox.Height / 11)-Step(ScaleBox.Height / 11, ScaleBox.Height / 11), RGB(0, 0, shade(CSng(i), 0, 11)), BF
        End Select
        
        ScaleBox.AutoRedraw = False
    Next i
    
End Sub
' User has clicked on map. Find closest creature and display stats.
Private Sub MapBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim c As Integer
    Dim cx As Single
    Dim cy As Single
    Dim nearest As Single
    Dim dist As Single
    Dim critter As Integer
    
    ' convert click coordinates to world coordinates
    cx = x / MapBox.Width * (UBound(World, 1) + 1)
    cy = y / MapBox.Height * (UBound(World, 2) + 1)
    
    nearest = 9999
    For c = 0 To UBound(Life)
        dist = Sqr((cx - Life(c).x) ^ 2 + (cy - Life(c).y) ^ 2)
        If dist < nearest Then
            nearest = dist
            critter = c
        End If
    Next c

    CreatureDataForm.CreatureBox = critter
    UpdateCreatureForm (critter)
    CreatureDataForm.Show

End Sub

Private Sub UpdateCreatureForm(critter As Integer)

    CreatureDataForm.DataBox(0) = Life(critter).Age
    CreatureDataForm.DataBox(1) = Life(critter).Size
    CreatureDataForm.DataBox(2) = Life(critter).Food
    CreatureDataForm.DataBox(3) = Life(critter).EatFactor
    CreatureDataForm.DataBox(4) = Life(critter).Health
    CreatureDataForm.DataBox(5) = Life(critter).Metabolism
    CreatureDataForm.DataBox(6) = Life(critter).Species
    CreatureDataForm.DataBox(7) = Life(critter).XSpeed
    CreatureDataForm.DataBox(8) = Life(critter).XGoal
    CreatureDataForm.DataBox(9) = Life(critter).x
    CreatureDataForm.DataBox(10) = Life(critter).MaxAge
    CreatureDataForm.DataBox(11) = Life(critter).MaxSize
    CreatureDataForm.DataBox(12) = Life(critter).MaxFood
    CreatureDataForm.DataBox(13) = Life(critter).SizeToBreed
    CreatureDataForm.DataBox(14) = Life(critter).Injury
    CreatureDataForm.DataBox(15) = Life(critter).Rested
    CreatureDataForm.DataBox(16) = Life(critter).MaxSpeed
    CreatureDataForm.DataBox(17) = Life(critter).YSpeed
    CreatureDataForm.DataBox(18) = Life(critter).YGoal
    CreatureDataForm.DataBox(19) = Life(critter).y
    CreatureDataForm.DataBox(20) = Life(critter).GrazeFactor
    
    
End Sub

Private Sub MapModeBtn_Click(Index As Integer)

    MapMode = Index
    Select Case Index
        Case 0:             ' look at food
            ScaleLo = 0
            ScaleHi = 25
        Case 1:             ' look at fertility
            ScaleLo = 2
            ScaleHi = 7
    End Select
    DrawMapScale
    
End Sub

Private Sub PhysicsBtn_Click()

    PhysicsForm.Show
    
End Sub

Private Sub RandBtn_Click()

    Randomize
    
End Sub

Private Sub RestartBtn_Click()
    
    MakeWorld           ' Set terrain constants and variables
    MakeCreatures       ' Populate the world
    Ticks = 0           ' time = 0
    GuiInit
    ClockBox.Text = 0
    Timer1.Enabled = True

End Sub

Private Sub RunFastBtn_Click()
    
    RunFast = Not RunFast
    Timer1.Enabled = True
    If RunFast Then
        MessageBox.Text = "Running at high speed"
    Else
        MessageBox.Text = "Running at regular speed"
    End If
    
End Sub

Private Sub SpeedBtn_Click(Index As Integer)

    Select Case Index
        Case 0:
            Every = 1
        Case 1:
            Every = 10
        Case 2:
            Every = 50
        Case 3:
            Every = 100
        Case 4:
            Every = 500
    End Select
            
End Sub

Private Sub StartBtn_Click()

    Timer1.Enabled = True
    If RunFast Then
        MessageBox.Text = "Running at high speed"
    Else
        MessageBox.Text = "Running at regular speed"
    End If

End Sub

Private Sub StopBtn_Click()

    Timer1.Enabled = False
    MessageBox.Text = "Paused"

End Sub

Private Sub Timer1_Timer()

    Dim c As Integer
    Dim cycles As Integer
    
    cycles = Every
    
    While cycles > 0
        GrowStuff
        
        For c = 0 To UBound(Life)
            Call ProcessCreature(c)
        Next c
        
        Ticks = Ticks + 1
        cycles = cycles - 1
        DoEvents
    
        deathcounter = (deathcounter + 1) Mod UBound(AgeDeaths)
        AgeDeaths(deathcounter) = 0
        StarveDeaths(deathcounter) = 0
        ExhaustionDeaths(deathcounter) = 0
        
    Wend
    
    Call DrawWorldData(0, 0)
    Call DrawCreatureData(0)
    Call ShowStats

End Sub
Private Sub ShowStats()

    Dim i As Integer
    Dim metcount As Double
    Dim liveones As Double
    Dim bsizecount As Double
    Dim efcount As Double
    Dim gfcount As Double
    
    For i = 0 To UBound(Life)
        If Life(i).Health > 0 Then
            bsizecount = bsizecount + Life(i).SizeToBreed
            efcount = efcount + Life(i).EatFactor
            metcount = metcount + Life(i).Metabolism
            gfcount = gfcount + Life(i).GrazeFactor
            liveones = liveones + 1
        End If
    Next i
    
    ClockBox = Ticks
    AgeBox = AgeCounter
    StarveBox = StarveCounter
    ExhaustionBox = ExhaustionCounter
    LiveBox = liveones
    
    MetBox.Text = Format((metcount / liveones), "###.0000")
    ChartData(4, chartpoint) = (metcount / liveones) * 50
    
    BAgeBox.Text = Format((efcount / liveones), "###.0000")
    ChartData(5, chartpoint) = (efcount / liveones) * 100
    
    BSizeBox.Text = Format((bsizecount / liveones), "###.0000")
    ChartData(6, chartpoint) = (bsizecount / liveones) - 50
    
    GFBox.Text = Format((gfcount / liveones), "###.0000")
    ChartData(7, chartpoint) = (gfcount / liveones) * 100
    
End Sub
Private Sub GrowStuff()

    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To UBound(World, 1)
        For j = 0 To UBound(World, 2)
            World(i, j).Biomass = World(i, j).Biomass + World(i, j).Fertility * FertilityFactor
            
            ' limit biomass to 5 times fertility
            If World(i, j).Biomass > World(i, j).Fertility * BiomassMaxFactor Then
                World(i, j).Biomass = World(i, j).Fertility * BiomassMaxFactor
            End If
            
        Next j
    Next i
    
End Sub
' Draw the world. At this point, color scale shows available food
'
Private Sub DrawWorldData(tlx As Integer, tly As Integer)

    Dim i As Integer
    Dim j As Integer
    Dim dx As Single
    Dim dy As Single
    Dim sx As Single
    Dim sy As Single
    Dim ccolor As Long
    Dim clr1 As Integer
    Dim clr2 As Integer
    Dim slo As Single
    Dim shi As Single
    
    sx = UBound(World, 1) - LBound(World, 1) + 1
    sy = UBound(World, 2) - LBound(World, 2) + 1
    dx = MapBox.Width / sx
    dy = MapBox.Height / sy
    
    ' Drap map box
    slo = ScaleLo
    shi = ScaleHi
    Select Case MapMode
        Case 0:             ' Food
            For i = 0 To UBound(World, 1)
                For j = 0 To UBound(World, 2)
                    MapBox.Line (i * dx, j * dy)-Step(dx, dy), RGB(0, shade(World(i, j).Biomass, slo, shi), 0), BF
                Next j
            Next i
         Case 1:             ' Fertilty
            For i = 0 To UBound(World, 1)
                For j = 0 To UBound(World, 2)
                    MapBox.Line (i * dx, j * dy)-Step(dx, dy), RGB(0, 0, shade(World(i, j).Fertility, slo, shi)), BF
                Next j
            Next i
   End Select
   
    ' Draw little white dots for live critters, grey for dead
    ' get color transitions from gui
    clr1 = ClrBox(0)
    clr2 = ClrBox(1)
    
    For i = 0 To UBound(Life)
        If Life(i).Health > 0 Then
            ccolor = RGB(255, 0, 0)
            ' If young adult, then white
            If Life(i).Size > clr1 Then ccolor = RGB(255, 255, 255)
            If Life(i).Size >= clr2 Then ccolor = RGB(255, 255, 0)
        Else
            ' Dead one...
            ccolor = RGB(128, 128, 128)
        End If
        MapBox.Line (Life(i).x * dx - 20, Life(i).y * dy - 20)-Step(40, 40), ccolor, BF
    Next i
    
End Sub
' Draw creature date for the first n creatures starting with creature sc
'
Private Sub DrawCreatureData(sc As Integer)

    Dim i As Integer
    Dim deaths As Integer
    Dim A As Integer
    Dim s As Integer
    Dim e As Integer
    
    
    'ShowStats
    
    If CreatureDataForm.Visible Then
       UpdateCreatureForm (Val(CreatureDataForm.CreatureBox))
    End If
    
    ' Draw graph
    For i = 1 To UBound(AgeDeaths)
        A = A + AgeDeaths(i)
        s = s + StarveDeaths(i)
        e = e + ExhaustionDeaths(i)
    Next i
    
    deaths = A + s + e
    
    If deaths = 0 Then Exit Sub
        
    ChartData(0, chartpoint) = A / deaths * 100
    ChartData(1, chartpoint) = s / deaths * 100
    ChartData(2, chartpoint) = e / deaths * 100
    ChartData(3, chartpoint) = LiveCount / UBound(Life) * 100
    
    ChartBox.Cls
    
    ' Do population and death stats
    ChartBox.PSet (0, ChartData(0, (1 + chartpoint) Mod 200)), RGB(255, 0, 0)
    For i = 1 To 200
        ChartBox.Line -(i - 1, ChartData(0, (i + chartpoint) Mod 200)), RGB(255, 0, 0)
    Next i
    
    ChartBox.PSet (0, ChartData(1, (1 + chartpoint) Mod 200)), RGB(255, 255, 0)
    For i = 1 To 200
        ChartBox.Line -(i - 1, ChartData(1, (i + chartpoint) Mod 200)), RGB(255, 255, 0)
    Next i
    
    ChartBox.PSet (0, ChartData(2, (1 + chartpoint) Mod 200)), RGB(0, 255, 0)
    For i = 1 To 200
        ChartBox.Line -(i - 1, ChartData(2, (i + chartpoint) Mod 200)), RGB(0, 255, 0)
    Next i
    
    ChartBox.PSet (0, ChartData(3, (1 + chartpoint) Mod 200)), RGB(0, 0, 255)
    For i = 1 To 200
        ChartBox.Line -(i - 1, ChartData(3, (i + chartpoint) Mod 200)), RGB(0, 0, 255)
    Next i
    
    ' Do genetics stats
    GeneBox.Cls
    GeneBox.PSet (0, ChartData(4, (1 + chartpoint) Mod 200)), RGB(255, 0, 0)
    For i = 1 To 200
        GeneBox.Line -(i - 1, ChartData(4, (i + chartpoint) Mod 200)), RGB(255, 0, 0)
    Next i
    
    GeneBox.PSet (0, ChartData(5, (1 + chartpoint) Mod 200)), RGB(255, 255, 0)
    For i = 1 To 200
        GeneBox.Line -(i - 1, ChartData(5, (i + chartpoint) Mod 200)), RGB(255, 255, 0)
    Next i
    
    GeneBox.PSet (0, ChartData(6, (1 + chartpoint) Mod 200)), RGB(0, 255, 0)
    For i = 1 To 200
        GeneBox.Line -(i - 1, ChartData(6, (i + chartpoint) Mod 200)), RGB(0, 255, 0)
    Next i
    
    GeneBox.PSet (0, ChartData(7, (1 + chartpoint) Mod 200)), RGB(0, 255, 0)
    For i = 1 To 200
        GeneBox.Line -(i - 1, ChartData(7, (i + chartpoint) Mod 200)), RGB(0, 0, 255)
    Next i
    
    
    chartpoint = (chartpoint + 1) Mod 200
    
End Sub
' Do goal processing for creature c
'
Private Sub ProcessCreature(c As Integer)

    Dim resting As Boolean
    Dim ChoiceMade As Boolean
    Dim goal As Integer
    Dim motivation As Single
    Dim x As Single
    
    ' Don't bother with dead ones.
    If Life(c).Health = 0 Then
        Exit Sub
    End If
    
    goal = 1                ' Rest
    motivation = 0          ' no motivation
    
    ChoiceMade = False
    resting = False
    
    ' Eating once doesn't count as a choice.
    Call eat(c, Life(c).x, Life(c).y)
    
    motivation = (1 - (Life(c).Food / Life(c).MaxFood)) * Life(c).EatFactor
    goal = 1                ' eat
    
    x = (1 - (Life(c).Food / Life(c).MaxFood)) * Life(c).GrazeFactor
    If x > motivation Then
        motivation = x
        goal = 2            ' Look for food
    End If
    
    x = (1 - (Life(c).Rested / 10)) * 0.5
    If x > motivation Then
        motivation = x
        goal = 3            ' rest
    End If
    
    If (Life(c).Size >= Life(c).SizeToBreed) And motivation < 0.3 Then
        motivation = 0.3
        goal = 4            ' look for mate
    End If
    
    Select Case goal
        Case 1:             ' Still hungry - eat
            Call eat(c, Life(c).x, Life(c).y)
        Case 2:             ' Look for more food elsewhere
            ' Still hungry, look for better grazing.
            Call LookForFood(c)
            ' If there's more food elsewhere, start moving
            If Life(c).x <> Life(c).XGoal Or Life(c).y <> Life(c).YGoal Then
                Call MoveCreature(c)
            Else
                ' Eat again where we are.
                Call eat(c, Life(c).x, Life(c).y)
            End If
        Case 3:             ' We need to rest
            resting = True
        Case 4:             ' Look for mate
            Call LookForMate(c)
            If Life(c).x <> Life(c).XGoal Or Life(c).y <> Life(c).YGoal Then
                Call MoveCreature(c)
            End If
    End Select
    
    Call Metabolize(c, resting)
    
    If LiveCount <= (UBound(Life) / 10) Then
        StopBtn_Click
        MessageBox.Text = "Population has crashed. Try again"
    End If
    ' Another tick older and deeper in debt..
    Life(c).Age = Life(c).Age + 1
    
End Sub
Private Sub breed(c1 As Integer, c2 As Integer)

    Dim i As Integer
    
    ' Find a life slot (recycle a dead one)
    For i = 0 To UBound(Life)
        If Life(i).Health = 0 Then
            Exit For
        End If
    Next i
    
    ' No dead ones? error and exit
    If i >= UBound(Life) Then
        'MsgBox "Can't make new life - no slots"
        Exit Sub
    End If
        
    Life(i).Species = 1     ' all the same - we don't care
    Life(i).MaxSize = 100   ' 100 kgs
    Life(i).Size = 10       ' baby
    Life(i).Age = 1         ' 10 ticks
    If Rnd > 0.5 Then
        Life(i).EatFactor = 0.95 * Life(c1).EatFactor + Rnd * 0.1 ' How hungry before we eat?
    Else
        Life(i).EatFactor = 0.95 * Life(c2).EatFactor + Rnd * 0.1
    End If
    
    If Rnd > 0.5 Then
        Life(i).GrazeFactor = 0.95 * Life(c1).GrazeFactor + Rnd * 0.1 ' How hungry before we move on?
    Else
        Life(i).GrazeFactor = 0.95 * Life(c2).GrazeFactor + Rnd * 0.1
    End If
    
    If Rnd > 0.5 Then
        Life(i).SizeToBreed = (0.95 + Rnd * 0.1) * Life(c1).SizeToBreed        ' How heavy to breed?
    Else
        Life(i).SizeToBreed = (0.95 + Rnd * 0.1) * Life(c2).SizeToBreed       ' How heavy to breed?
    End If
    Life(i).MaxFood = 10    ' Can hold 10 kgs
    Life(i).Food = 2        ' a little hungry
    Life(i).Injury = 0      ' not injured
    Life(i).Rested = 10     ' rested
    ' mutate a little
    If Rnd > 0.5 Then
        Life(i).Metabolism = (0.95 + Rnd * 0.1) * Life(c1).Metabolism
    Else
        Life(i).Metabolism = (0.95 + Rnd * 0.1) * Life(c2).Metabolism
    End If
    Life(i).Health = 10     ' healthy
    Life(i).XSpeed = 0       ' Not moving
    Life(i).YSpeed = 0       ' Not moving
    Life(i).x = Life(c1).x
    Life(i).y = Life(c1).y
    Life(i).XGoal = Life(i).x
    Life(i).YGoal = Life(i).y
    
    Constrain (i)
    
    ' Cost to parents
    Life(c1).Food = Life(c1).Food / 2
    Life(c1).Rested = Life(c1).Rested / 2
    Life(c1).Size = Life(c1).Size - 15
    Life(c2).Food = Life(c2).Food / 2
    Life(c2).Rested = Life(c2).Rested / 2
    Life(c2).Size = Life(c2).Size - 15
    
    'MessageBox.Text = MessageBox.Text & "New creature " & i & " born to " & c1 & " and " & c2 & vbCrLf
    'MapBox.Circle (Life(c1).x, Life(c1).y), 15, RGB(255, 0, 0)
    LiveCount = LiveCount + 1

End Sub
' Critter c needs to process metabolic functions
'
Private Sub Metabolize(c As Integer, resting As Boolean)

    ' Gain some weight if our belly is full
    If Life(c).Food > 0.8 * Life(c).MaxFood Then
        Life(c).Size = Life(c).Size + 0.025 * (Life(c).MaxSize - Life(c).Size)
    End If
    
    If resting Then
        ' resting critter consumes .15 food and gains 5 rest
        Life(c).Rested = Life(c).Rested + 5
        Life(c).Food = Life(c).Food - 0.5 * HungerFactor * Life(c).Metabolism
    Else
        ' non-resting critter consumes .3 food and loses .25 rest
        Life(c).Rested = Life(c).Rested - 0.25
        Life(c).Food = Life(c).Food - HungerFactor * Life(c).Metabolism
    End If
    
    ' Did we starve to death?
    If Life(c).Food <= 0 Then
        Call DeathReport(c, "Starvation")
        Exit Sub
    End If
    
    ' Did we die of exhaustion?
    If Life(c).Rested <= 0 Then
        Call DeathReport(c, "Exhaustion")
        Exit Sub
    End If
     
     ' Did we die of old age?
    If Life(c).Age > Life(c).MaxAge Then
        Call DeathReport(c, "Old Age")
        Exit Sub
    End If
    
End Sub
Sub DeathReport(c As Integer, reason As String)
    
    Life(c).Food = 0
    Life(c).Health = 0
    If reason = "Old Age" Then
        AgeCounter = AgeCounter + 1
        AgeDeaths(deathcounter) = AgeDeaths(deathcounter) + 1
    End If
    
    If reason = "Starvation" Then
        StarveCounter = StarveCounter + 1
        StarveDeaths(deathcounter) = StarveDeaths(deathcounter) + 1
    End If
        
    If reason = "Exhaustion" Then
        ExhaustionCounter = ExhaustionCounter + 1
        ExhaustionDeaths(deathcounter) = ExhaustionDeaths(deathcounter) + 1
    End If
    
    'MessageBox.Text = MessageBox.Text & "Creature " & c & " died of " & reason
    'MessageBox.Text = MessageBox.Text & " Met = " & Life(c).Metabolism & vbCrLf
    LiveCount = LiveCount - 1

End Sub
' Look around for adjacent cells with more food than here.
' If one is found, set a course in that direction
'
Private Sub LookForFood(c As Integer)

    Dim x As Integer
    Dim x2 As Integer
    Dim y As Integer
    Dim lx As Integer
    Dim ly As Integer
    Dim rx As Integer
    Dim ry As Integer
    Dim i As Integer
    Dim j As Integer
    Dim BestFood As Single
    Dim bestx As Single
    Dim besty As Single
    
    ' where are we now?
    x = Int(Life(c).x)
    y = Int(Life(c).y)
    BestFood = World(x, y).Biomass
    bestx = x
    besty = y
      
    ' What are our search boundaries?
    ' We'd like to search +/- 1 cell in all directions
    lx = x - 1
    rx = x + 1
    ly = y - 1
    ry = y + 1
    ' But edges may be in the way
    'If lx < 0 Then lx = 0
    'If rx > UBound(World, 1) Then rx = UBound(World, 1)
    
    If ly < 0 Then ly = 0
    If ry > UBound(World, 2) Then ry = UBound(World, 2)
    
    ' Search each cell in our range of sensing
     For i = lx To rx
        x2 = (i + UBound(World, 1) + 1) Mod (UBound(World, 1) + 1)
        For j = ly To ry
            ' Look at apparent food in cell i,j
            If SeeFood(c, x2, j) > BestFood Then
                BestFood = World(x2, j).Biomass
                bestx = x2
                besty = j
            End If
        Next j
    Next i
    
    If c = 0 Then
        DoEvents
    End If
    
    ' Bestx and besty contain best cell within view. Is it somewhere else?
    If bestx <> x Or besty <> y Then
        ' There are greener pastures... Are we already heading somewhere?
        If Int(Life(c).XGoal) <> bestx Or Int(Life(c).YGoal) <> besty Then
            ' We're already headed elsewhere. Is there enough reason to change course?
            If SeeFood(c, Int(Life(c).XGoal), Int(Life(c).YGoal)) < (0.8 * BestFood) Then
                ' If so, head there
                bestx = bestx + Rnd()
                besty = besty + Rnd()
                Life(c).XGoal = bestx
                Life(c).YGoal = besty
            End If
        End If
    End If
    
End Sub
' Look around for Mate.
' If one is found, set a course in that direction
'
Private Sub LookForMate(c As Integer)

    Dim x As Integer
    Dim y As Integer
    Dim i As Integer
    Dim dist
    Dim BestDist As Single
    Dim bestmate As Integer
    Dim bestx As Single
    Dim besty As Single
    
    ' where are we now?
    x = Int(Life(c).x)
    y = Int(Life(c).y)
    BestDist = 9999
    bestx = x
    besty = y
    
    ' Search each creature
    For i = 0 To UBound(Life)
        ' We're not necrophilic - avoid dead ones
        If Life(i).Health = 0 Then GoTo nexti
        ' We're not hermaphroditic - don't try to mate with self
        If i = c Then GoTo nexti
        ' Reject creatures who cannot breed
        If Life(i).Size < Life(i).SizeToBreed Then GoTo nexti
        
        ' This one can breed - how far away is it?
        dist = Sqr((Life(i).x - x) ^ 2 + (Life(i).y - y) ^ 2)
        ' only criteria at this time is proximity - we'll mate with most convenient.
        If dist < BestDist Then
            bestmate = i
            BestDist = dist
            bestx = Life(i).x
            besty = Life(i).y
        End If
nexti:
    Next i
    
    ' exit if no mates
    If BestDist = 9999 Then Exit Sub
    
    If BestDist < 0.5 Then
        ' We're mated - let's breed
        Call breed(c, bestmate)
    Else
        ' Let's go - we're lustful!
        Life(c).XGoal = bestx
        Life(c).YGoal = besty
    End If
    
End Sub
Private Sub MoveCreature(c As Integer)
        
    Dim dx As Single
    Dim dy As Single
        
   If c = 0 Then
        DoEvents
    End If
    
    dx = Life(c).XGoal - Life(c).x
    ' if it would be closer to wrap around:
    If Abs(dx) > UBound(World, 1) / 2 Then
        dx = Fmod((Life(c).XGoal + UBound(World, 1) / 2), UBound(World, 1) + 1) - Fmod((Life(c).x + UBound(World, 1) / 2), UBound(World, 1) + 1)
    End If
    
    dy = Life(c).YGoal - Life(c).y
    
    ' Full speed for now.
    Life(c).XSpeed = Life(c).MaxSpeed * dx / (Abs(dx) + Abs(dy))
    Life(c).YSpeed = Life(c).MaxSpeed * dy / (Abs(dx) + Abs(dy))

    ' Have we reached our goal in the x axis?
    ' If not, then move. Our speed is .01 cells per tick per speed unit.
    If Abs(dx) < Abs(Life(c).XSpeed * 0.01) Then
        Life(c).x = Life(c).XGoal
        Life(c).XSpeed = 0
    Else
        ' Allow X wraparound
        Life(c).x = Fmod((Life(c).x + Life(c).XSpeed * 0.01), (UBound(World, 1) + 1))
        'Life(c).x = Life(c).x + Life(c).XSpeed * 0.01
    End If
        
    ' Have we reached our goal in the y axis?
    If Abs(dy) < Abs(Life(c).YSpeed * 0.01) Then
        Life(c).y = Life(c).YGoal
        Life(c).YSpeed = 0
    Else
        Life(c).y = Life(c).y + Life(c).YSpeed * 0.01
    End If
        
End Sub

' Creature c, who is at coordinates wx,wy wants to eat.
' Eventually, we'll have to deal with multiple critters in same cell
'
Private Sub eat(c As Integer, wx As Single, wy As Single)

    Dim Meal As Integer

    ' How much food is there in this cell? We can't eat more than 10%
    Meal = World(Int(wx), Int(wy)).Biomass * 0.1

    ' Can this critter eat it all?
    If Life(c).MaxFood - Life(c).Food < Meal Then
        'if not, reduce meal size
        Meal = Life(c).MaxFood - Life(c).Food
    End If
    
    ' Critter gets fuller belly
    Life(c).Food = Life(c).Food + Meal
    
    ' Available food is diminished
    World(Int(wx), Int(wy)).Biomass = World(Int(wx), Int(wy)).Biomass - Meal

End Sub
' Critter c attempts to see (with imperfect vision) how much food is in cell x,y
'
Private Function SeeFood(c As Integer, x As Integer, y As Integer) As Single

    ' sensed food is correct +/- 10%
    SeeFood = (World(x, y).Biomass * (0.9 + Rnd() * 0.2))

End Function

Function Fmod(A As Single, B As Single) As Single

    While A < 0
        A = A + B
    Wend

    While A > B
        A = A - B
    Wend
        
    Fmod = A

End Function
