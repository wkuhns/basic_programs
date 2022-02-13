VERSION 5.00
Begin VB.Form PhysicsForm 
   Caption         =   "Physics"
   ClientHeight    =   4572
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   8544
   LinkTopic       =   "Form1"
   ScaleHeight     =   4572
   ScaleWidth      =   8544
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox WorldBox 
      Height          =   288
      Index           =   0
      Left            =   7440
      TabIndex        =   23
      Top             =   2880
      Width           =   732
   End
   Begin VB.TextBox WorldBox 
      Height          =   288
      Index           =   1
      Left            =   7440
      TabIndex        =   22
      Top             =   3240
      Width           =   732
   End
   Begin VB.TextBox WorldBox 
      Height          =   288
      Index           =   2
      Left            =   7440
      TabIndex        =   21
      Top             =   3600
      Width           =   732
   End
   Begin VB.TextBox WorldBox 
      Height          =   288
      Index           =   3
      Left            =   7440
      TabIndex        =   20
      Top             =   3960
      Width           =   732
   End
   Begin VB.TextBox GeneBox 
      Height          =   288
      Index           =   3
      Left            =   3240
      TabIndex        =   17
      Top             =   3960
      Width           =   732
   End
   Begin VB.TextBox GeneBox 
      Height          =   288
      Index           =   2
      Left            =   3240
      TabIndex        =   15
      Top             =   3600
      Width           =   732
   End
   Begin VB.TextBox GeneBox 
      Height          =   288
      Index           =   1
      Left            =   3240
      TabIndex        =   13
      Top             =   3240
      Width           =   732
   End
   Begin VB.TextBox GeneBox 
      Height          =   288
      Index           =   0
      Left            =   3240
      TabIndex        =   11
      Top             =   2880
      Width           =   732
   End
   Begin VB.CommandButton SetBtn 
      Caption         =   "Set Constants and exit"
      Height          =   312
      Left            =   5820
      TabIndex        =   9
      Top             =   60
      Width           =   2412
   End
   Begin VB.TextBox FertilityBox 
      Height          =   288
      Index           =   3
      Left            =   3240
      TabIndex        =   8
      Top             =   1920
      Width           =   732
   End
   Begin VB.TextBox FertilityBox 
      Height          =   288
      Index           =   2
      Left            =   3240
      TabIndex        =   6
      Top             =   1560
      Width           =   732
   End
   Begin VB.TextBox FertilityBox 
      Height          =   288
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   1200
      Width           =   732
   End
   Begin VB.TextBox FertilityBox 
      Height          =   288
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   732
   End
   Begin VB.Label Label2 
      Caption         =   "World size (cells per axis)"
      Height          =   252
      Index           =   11
      Left            =   4560
      TabIndex        =   28
      Top             =   2880
      Width           =   2772
   End
   Begin VB.Label Label2 
      Caption         =   "Maximum creature population"
      Height          =   252
      Index           =   10
      Left            =   4560
      TabIndex        =   27
      Top             =   3240
      Width           =   2772
   End
   Begin VB.Label Label2 
      Caption         =   "(Reserved for future use)"
      Height          =   252
      Index           =   9
      Left            =   4560
      TabIndex        =   26
      Top             =   3600
      Width           =   2772
   End
   Begin VB.Label Label2 
      Caption         =   "(Reserved for future use)"
      Height          =   252
      Index           =   8
      Left            =   4560
      TabIndex        =   25
      Top             =   3960
      Width           =   2772
   End
   Begin VB.Shape Shape1 
      Height          =   1932
      Index           =   2
      Left            =   4320
      Top             =   2400
      Width           =   3972
   End
   Begin VB.Label Label3 
      Caption         =   "World Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   4440
      TabIndex        =   24
      Top             =   2520
      Width           =   2292
   End
   Begin VB.Label Label3 
      Caption         =   "Genetic constraints"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   19
      Top             =   2520
      Width           =   2292
   End
   Begin VB.Shape Shape1 
      Height          =   1932
      Index           =   1
      Left            =   120
      Top             =   2400
      Width           =   3972
   End
   Begin VB.Label Label2 
      Caption         =   "Metabolism to food consumption factor"
      Height          =   252
      Index           =   7
      Left            =   360
      TabIndex        =   18
      Top             =   3960
      Width           =   2772
   End
   Begin VB.Label Label2 
      Caption         =   "Metabolism to max speed factor"
      Height          =   252
      Index           =   6
      Left            =   360
      TabIndex        =   16
      Top             =   3600
      Width           =   2772
   End
   Begin VB.Label Label2 
      Caption         =   "Metabolism to Max Age dividened"
      Height          =   252
      Index           =   5
      Left            =   360
      TabIndex        =   14
      Top             =   3240
      Width           =   2772
   End
   Begin VB.Label Label2 
      Caption         =   "Mutation factor"
      Height          =   252
      Index           =   4
      Left            =   360
      TabIndex        =   12
      Top             =   2880
      Width           =   2772
   End
   Begin VB.Shape Shape1 
      Height          =   1812
      Index           =   0
      Left            =   120
      Top             =   480
      Width           =   3972
   End
   Begin VB.Label Label3 
      Caption         =   "Biomass and Fertility"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   2532
   End
   Begin VB.Label Label2 
      Caption         =   "Maximum biomass as mult. of fertility"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   2772
   End
   Begin VB.Label Label2 
      Caption         =   "Minimum possible fertility"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2772
   End
   Begin VB.Label Label2 
      Caption         =   "Maximum possible fertility"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2772
   End
   Begin VB.Label Label2 
      Caption         =   "Biomass production per unit of fertility"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Physical Constants"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2160
      TabIndex        =   0
      Top             =   60
      Width           =   3672
   End
End
Attribute VB_Name = "PhysicsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    FertilityBox(0) = FertilityFactor
    FertilityBox(1) = FertilityMax
    FertilityBox(2) = FertilityMin
    FertilityBox(3) = BiomassMaxFactor

    GeneBox(0) = Mut
    GeneBox(1) = MetAgeFactor
    GeneBox(2) = SpeedFactor
    GeneBox(3) = HungerFactor
    
    WorldBox(0) = WorldSize
    WorldBox(1) = LifeSize

End Sub

Private Sub SetBtn_Click()
    
    FertilityFactor = Val(FertilityBox(0))
    FertilityMax = Val(FertilityBox(1))
    FertilityMin = Val(FertilityBox(2))
    BiomassMaxFactor = Val(FertilityBox(3))

    Mut = Val(GeneBox(0))
    MetAgeFactor = Val(GeneBox(1))
    SpeedFactor = Val(GeneBox(2))
    HungerFactor = Val(GeneBox(3))
    
    WorldSize = Val(WorldBox(0))
    LifeSize = Val(WorldBox(1))
    
    PhysicsForm.Hide
    
End Sub

