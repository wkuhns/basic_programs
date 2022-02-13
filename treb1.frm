VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Data"
   ClientHeight    =   4656
   ClientLeft      =   720
   ClientTop       =   1272
   ClientWidth     =   7140
   Height          =   4980
   Left            =   672
   LinkTopic       =   "Form1"
   ScaleHeight     =   4656
   ScaleWidth      =   7140
   Top             =   996
   Width           =   7236
   Begin VB.CommandButton StopBtn 
      Caption         =   "Stop"
      Height          =   372
      Left            =   6120
      TabIndex        =   49
      Top             =   1320
      Width           =   972
   End
   Begin VB.CommandButton StartBtn 
      Caption         =   "Start"
      Height          =   372
      Left            =   6120
      TabIndex        =   48
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox Results 
      Height          =   288
      Index           =   11
      Left            =   2880
      TabIndex        =   46
      Top             =   4200
      Width           =   972
   End
   Begin VB.TextBox Results 
      Height          =   288
      Index           =   10
      Left            =   2880
      TabIndex        =   44
      Top             =   3840
      Width           =   972
   End
   Begin VB.TextBox Results 
      Height          =   288
      Index           =   9
      Left            =   2880
      TabIndex        =   42
      Top             =   3480
      Width           =   972
   End
   Begin VB.TextBox Results 
      Height          =   288
      Index           =   8
      Left            =   2880
      TabIndex        =   40
      Top             =   3120
      Width           =   972
   End
   Begin VB.TextBox Results 
      Height          =   288
      Index           =   7
      Left            =   2880
      TabIndex        =   38
      Top             =   2760
      Width           =   972
   End
   Begin VB.TextBox Results 
      Height          =   288
      Index           =   6
      Left            =   2880
      TabIndex        =   36
      Top             =   2400
      Width           =   972
   End
   Begin VB.TextBox Results 
      Height          =   288
      Index           =   5
      Left            =   2880
      TabIndex        =   34
      Top             =   2040
      Width           =   972
   End
   Begin VB.TextBox Results 
      Height          =   288
      Index           =   4
      Left            =   2880
      TabIndex        =   32
      Top             =   1680
      Width           =   972
   End
   Begin VB.TextBox Results 
      Height          =   288
      Index           =   3
      Left            =   2880
      TabIndex        =   30
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox Results 
      Height          =   288
      Index           =   2
      Left            =   2880
      TabIndex        =   28
      Top             =   960
      Width           =   972
   End
   Begin VB.TextBox Results 
      Height          =   288
      Index           =   1
      Left            =   2880
      TabIndex        =   26
      Top             =   600
      Width           =   972
   End
   Begin VB.TextBox Results 
      Height          =   288
      Index           =   0
      Left            =   2880
      TabIndex        =   24
      Top             =   240
      Width           =   972
   End
   Begin VB.TextBox ParamBox 
      Height          =   288
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   4200
      Width           =   972
   End
   Begin VB.TextBox ParamBox 
      Height          =   288
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Top             =   3840
      Width           =   972
   End
   Begin VB.TextBox ParamBox 
      Height          =   288
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   972
   End
   Begin VB.TextBox ParamBox 
      Height          =   288
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   972
   End
   Begin VB.TextBox ParamBox 
      Height          =   288
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Text            =   "18"
      Top             =   2760
      Width           =   972
   End
   Begin VB.TextBox ParamBox 
      Height          =   288
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Text            =   "18"
      Top             =   2400
      Width           =   972
   End
   Begin VB.TextBox ParamBox 
      Height          =   288
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Text            =   "3.5"
      Top             =   2040
      Width           =   972
   End
   Begin VB.TextBox ParamBox 
      Height          =   288
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Text            =   "75"
      Top             =   1680
      Width           =   972
   End
   Begin VB.TextBox ParamBox 
      Height          =   288
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Text            =   "180"
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox ParamBox 
      Height          =   288
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Text            =   "315"
      Top             =   960
      Width           =   972
   End
   Begin VB.TextBox ParamBox 
      Height          =   288
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Text            =   ".05"
      Top             =   600
      Width           =   972
   End
   Begin VB.TextBox ParamBox 
      Height          =   288
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "3.0"
      Top             =   240
      Width           =   972
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6600
      Top             =   120
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   252
      Index           =   11
      Left            =   3960
      TabIndex        =   47
      Top             =   4200
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   252
      Index           =   10
      Left            =   3960
      TabIndex        =   45
      Top             =   3840
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   252
      Index           =   9
      Left            =   3960
      TabIndex        =   43
      Top             =   3480
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   252
      Index           =   8
      Left            =   3960
      TabIndex        =   41
      Top             =   3120
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   252
      Index           =   7
      Left            =   3960
      TabIndex        =   39
      Top             =   2760
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "6: Y position"
      Height          =   252
      Index           =   6
      Left            =   3960
      TabIndex        =   37
      Top             =   2400
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "5: X position"
      Height          =   252
      Index           =   5
      Left            =   3960
      TabIndex        =   35
      Top             =   2040
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "4: Y velocity"
      Height          =   252
      Index           =   4
      Left            =   3960
      TabIndex        =   33
      Top             =   1680
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "3: X Velocity"
      Height          =   252
      Index           =   3
      Left            =   3960
      TabIndex        =   31
      Top             =   1320
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "2: Arm rotation"
      Height          =   252
      Index           =   2
      Left            =   3960
      TabIndex        =   29
      Top             =   960
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "1: Y Force"
      Height          =   252
      Index           =   1
      Left            =   3960
      TabIndex        =   27
      Top             =   600
      Width           =   1572
   End
   Begin VB.Label Label2 
      Caption         =   "0: X Force"
      Height          =   252
      Index           =   0
      Left            =   3960
      TabIndex        =   25
      Top             =   240
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   " "
      Height          =   252
      Index           =   11
      Left            =   1200
      TabIndex        =   23
      Top             =   4200
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   " "
      Height          =   252
      Index           =   10
      Left            =   1200
      TabIndex        =   21
      Top             =   3840
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   " "
      Height          =   252
      Index           =   9
      Left            =   1200
      TabIndex        =   19
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   " "
      Height          =   252
      Index           =   8
      Left            =   1200
      TabIndex        =   17
      Top             =   3120
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "7: Length of sling "
      Height          =   252
      Index           =   7
      Left            =   1200
      TabIndex        =   15
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "6: Arm axis -> sling"
      Height          =   252
      Index           =   6
      Left            =   1200
      TabIndex        =   13
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "5: Arm axis -> weight"
      Height          =   252
      Index           =   5
      Left            =   1200
      TabIndex        =   11
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "4: Theta at release"
      Height          =   252
      Index           =   4
      Left            =   1200
      TabIndex        =   9
      Top             =   1680
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "3: Theta of sling"
      Height          =   252
      Index           =   3
      Left            =   1200
      TabIndex        =   7
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "2: Theta of arm"
      Height          =   252
      Index           =   2
      Left            =   1200
      TabIndex        =   5
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "1: Mass of projectile"
      Height          =   252
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "0: Mass of weight"
      Height          =   252
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   1500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Dim framecount As Integer
Dim energy As Double

Dim PA As Double        ' Polar moment of arm
Dim PW As Double        ' Polar moment of weight

Dim MW As Double        ' Mass of weight
Dim MP As Double        ' Mass of projectile

Dim TA As Double        ' Theta of arm
Dim TS As Double        ' Theta of sling
Dim TR As Double        ' Theta of arm at release

Dim QW As Double        ' Torque on arm due to weight
Dim QS As Double        ' Torque on arm due to sling

Dim FS As Double        ' Force on sling
Dim FPX As Double       ' Force on Projectile, X
Dim FPY As Double       ' Force on Projectile, Y

Dim LAW As Double       ' Length from axis to weight
Dim LAS As Double       ' Length from axis to sling
Dim LS As Double        ' Length of sling

Dim AAR As Double       ' Acceleration of Arm, radial
Dim APX As Double       ' Acceleration of Projectile, X
Dim APY As Double       ' Acceleration of Projectile, Y
Dim AWY As Double       ' Acceleration of Weight, Y

Dim VAR As Double       ' Velocity of Arm, radial
Dim VPX As Double       ' Velocity of Projectile, X
Dim VPY As Double       ' Velocity of Projectile, Y
Dim VWY As Double       ' Velocity of Weight, Y

Dim XP As Double        ' X position of projectile
Dim YP As Double        ' Y position of projectile
Dim YW As Double        ' Y position of Weight

Dim XA As Double        ' X position of end of arm
Dim YA As Double        ' Y position of end of arm

Const delta = 0.001

Sub DoFrame()

Dim theta As Double
Dim velocity As Double
Dim position As Double
Dim retard As Double
Dim offset As Double
Dim newVPX As Double
Dim newVPY As Double
Dim newXP As Double
Dim newYP As Double
Dim effectivemass As Double
Dim NetTorque As Double
Dim QR As Double
Dim FBS As Double
Dim FBW As Double
Dim AP As Double

' Torque balance: All forces that induce torque
' must cancel out. Forces are static loads and forces
' that produce acceleration of masses.

' FBW is force on beam due to weight
' FBS is force on beam due to sling
' We compute the equivalent force that would have to
' be connected at the 1" point of a horizontal beam
' to create the same torque.

    ' Effective force at attachment point
    FBW = (MW - (MW * 384 * AWY))
    
    ' Compensate for angle of beam
    FBW = FBW * Cos(TA)
    
    ' Effective mass at 1"
    FBW = FBW * LAW
    
    ' Arbitrary multiple of acceleration
    'QR = 0.1 * AAR
    
    FBS = (MP * Sin(TS) - (MP * 384 * AP)) * Sin(TS - TA)
    FBS = FBS * LAS
    
    ' Acceleration, velocity, and position of arm
    AAR = FBW / FBW + FBS
    VAR = VAR * AAR * delta
    TA = TA + VAR * delta
    
    ' position, velocity, and acceleration of weight
    YW = LAW * Sin(TA)
    VWY = VAR * Cos(TA) * LAW
    AWY = AAR * Cos(TA) * LAW

    ' Calculate X and Y at end of arm
    XA = LAS * Cos(TA)
    YA = LAS * Sin(TA)
    


'Calculate X and Y accelerations on projectile
    APX = 32 * FPX / MP
    APY = 32 * FPY / MP
    AP = Sqr(APX * APX + APY * APY)
    
'Calculate X and Y velocity of projectile
    newVPX = VPX + (APX * delta)
    newVPY = VPY + (APY * delta)
    
' Calculate new projectile position
    
    newXP = XP + VPX * delta
    newYP = YP + VPY * delta
    
    If Sqr((newXP - XA) * (newXP - XA) + (newYP - YA) * (newYP - YA)) > LS Then
    ' we're past perpendicular. Compute force on sling to correct
    ' projectile path.
        While Sqr((newXP - XA) * (newXP - XA) + (newYP - YA) * (newYP - YA)) > LS
            FS = FS * 1.001
            FPX = FS * Cos(TS)
            FPY = (FS * Sin(TS)) - MP
            APX = 32 * FPX / MP
            APY = 32 * FPY / MP
            newVPX = VPX + (APX * delta)
            newVPY = VPY + (APY * delta)
            newXP = XP + newVPX * delta
            newYP = YP + newVPY * delta
        Wend
    End If

    XP = newXP
    YP = newYP
    VPX = newVPX
    VPY = newVPY
    
'Calculate new arm position and acceleration

theta = TA
    
    While Sqr((XP - XA) * (XP - XA) + (YP - YA) * (YP - YA)) < LS
        TA = TA + 0.00001
        XA = LAS * Cos(TA)
        YA = LAS * Sin(TA)
    Wend

'Calculate weight position and acceleration. It may be that the
'projectile has modified the torque on the arm...

'FS = QW / (LAS * Sin(TS - TA))
'QW = FS * (LAS * Sin(TS - TA)


'Calculate force exerted by weight. In free-fall, weight would
'accelerate downwards at 384 in/sec/sec.
' f = ma        3 = (3/384)*384
'(fa - fg) = ma
'fa = ma + fg

retard = (MW / 384) * AWY + MW
QW = LAW * Cos(TA) * retard

'Calculate new sling angle

    TS = Atn((YA - YP) / (XA - XP))
    If XP < XA Then
        TS = TS + 3.14159265
    End If
    
If framecount Mod 50 = 0 Then

    Call PostResults

End If

framecount = framecount + 1
    
End Sub

Sub PostResults()

Dim wx As Single
Dim wy As Single
Dim sx As Single
Dim sy As Single
Dim px As Single
Dim py As Single

Results(0).Text = Format(FPX, "###.###")
Results(1).Text = Format(FPY, "###.###")
ParamBox(2) = Format(TA * 57.3, "###.###")
ParamBox(3) = Format(TS * 57.3, "###.###")
Results(2) = Format(AAR, "###.###")
Results(3) = Format(VPX, "###.###")
Results(4) = Format(VPY, "###.###")
Results(5) = Format(XP, "###.###")
Results(6) = Format(YP, "###.###")

'W is weight end of arm
'S is sling end of arm
'P is projectile

wx = 50 - (LAW * Cos(TA))
wy = 50 + (LAW * Sin(TA))
sx = 50 + (LAS * Cos(TA))
sy = 50 - (LAS * Sin(TA))
px = sx + (LS * Cos(TS))
py = sy - (LS * Sin(TS))

Form2.Picture1.Line (wx, wy)-(sx, sy), RGB(255, 0, 0)
Form2.Picture1.Line (sx, sy)-(px, py), RGB(0, 0, 255)

End Sub


Private Sub Command1_Click()

DoFrame

End Sub

Private Sub Form_Load()

    framecount = 0
    
    MW = ParamBox(0)
    MP = ParamBox(1)
    TA = ParamBox(2) / 57.3
    TS = ParamBox(3) / 57.3
    TR = ParamBox(4) / 57.3
    LAW = ParamBox(5)
    LAS = ParamBox(6)
    LS = ParamBox(7)

' Calculate the X and Y coordinates of the end of the arm.

    XA = LAS * Cos(TA)
    YA = LAS * Sin(TA)
    
' Calculate the X and Y coordinates of the projectile

    XP = LS * Cos(TS) + XA
    YP = LS * Sin(TS) + YA

'Calculate weight position and torque exerted

    YW = LAW * Sin(TA)
    QW = LAW * Cos(TA) * MW
     
' Calculate net energy of system in inch-pounds

energy = MW * 384 * YW
energy = energy + MP * 384 + YP

' Make other form appear

    Form2.Show
    
End Sub


Private Sub StartBtn_Click()
Timer1.Enabled = True
End Sub


Private Sub StopBtn_Click()

Timer1.Enabled = False

End Sub


Private Sub Timer1_Timer()

DoFrame

End Sub


