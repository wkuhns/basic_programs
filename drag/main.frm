VERSION 4.00
Begin VB.Form MainForm 
   Caption         =   "Drag Race Analysis"
   ClientHeight    =   7512
   ClientLeft      =   924
   ClientTop       =   1620
   ClientWidth     =   5796
   Height          =   7932
   Left            =   876
   LinkTopic       =   "Form1"
   ScaleHeight     =   7512
   ScaleWidth      =   5796
   Top             =   1248
   Width           =   5892
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Text            =   "10"
      Top             =   2400
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Text            =   "1"
      Top             =   2040
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Text            =   "4"
      Top             =   1680
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Text            =   ".9"
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Text            =   "1.96"
      Top             =   960
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Text            =   "2.1"
      Top             =   600
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "4"
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Maximum extension"
      Height          =   252
      Index           =   6
      Left            =   1200
      TabIndex        =   13
      Top             =   2400
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Mu"
      Height          =   252
      Index           =   5
      Left            =   1200
      TabIndex        =   11
      Top             =   2040
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Drive ratio"
      Height          =   252
      Index           =   4
      Left            =   1200
      TabIndex        =   9
      Top             =   1680
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Drivetrain efficiency"
      Height          =   252
      Index           =   3
      Left            =   1200
      TabIndex        =   7
      Top             =   1320
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Tire diameter"
      Height          =   252
      Index           =   2
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Mass"
      Height          =   252
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Number of springs"
      Height          =   252
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   1572
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'*******************************************************************
'
'   Program to calculate and plot drive wheel torques - see DRAG.DOC
'   for detailed documentation.
'
'*******************************************************************

    MAKE_TABLES
    
    nsprings = Val(MainForm.Text1(0).Text)
    mass = Val(MainForm.Text1(1).Text)
    tire_dia = Val(MainForm.Text1(2).Text)
    eff = Val(MainForm.Text1(3).Text)
    ratio = Val(MainForm.Text1(4).Text)
    mu = Val(MainForm.Text1(5).Text)
    maxext = Val(MainForm.Text1(6).Text)
    
    Open "logfile.dat" For Output As #2

'    For xx = 1 To pages
'        run$ = Str$(xx)
        INIT_TABLES
        GET_GOALS
        CALC_WINDS
        CALC_TORQUES
        GRAPH_TORQUES

        '*******************************************************************
        '   Format report header data
        '*******************************************************************

  '      ext$ = Str$(maxext) + "/" + Str$(actext)
        WRITE_GRAPH                       ' Write output to file
'    Next xx

End Sub


