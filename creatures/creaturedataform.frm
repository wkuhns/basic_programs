VERSION 5.00
Begin VB.Form CreatureDataForm 
   Caption         =   "Creature Data"
   ClientHeight    =   4740
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox CreatureBox 
      Height          =   288
      Left            =   2640
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   60
      Width           =   972
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   20
      Left            =   4680
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   4200
      Width           =   732
   End
   Begin VB.CommandButton CloseBtn 
      Caption         =   "Close"
      Height          =   372
      Left            =   4560
      TabIndex        =   40
      Top             =   60
      Width           =   972
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   19
      Left            =   4680
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   3840
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   18
      Left            =   4680
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   3480
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   17
      Left            =   4680
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   3120
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   16
      Left            =   4680
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   2760
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   15
      Left            =   4680
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   2400
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   14
      Left            =   4680
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   2040
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   13
      Left            =   4680
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   1680
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   12
      Left            =   4680
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   1320
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   11
      Left            =   4680
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   960
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   10
      Left            =   4680
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   600
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   9
      Left            =   1800
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3840
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   8
      Left            =   1800
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3480
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   7
      Left            =   1800
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3120
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   6
      Left            =   1800
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2760
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   5
      Left            =   1800
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2400
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   4
      Left            =   1800
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2040
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   3
      Left            =   1800
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1680
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1320
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   960
      Width           =   732
   End
   Begin VB.TextBox DataBox 
      Height          =   288
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   732
   End
   Begin VB.Shape Shape1 
      Height          =   4092
      Left            =   120
      Top             =   480
      Width           =   5412
   End
   Begin VB.Label Label2 
      Caption         =   "Creature #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   1260
      TabIndex        =   44
      Top             =   60
      Width           =   1272
   End
   Begin VB.Label Label1 
      Caption         =   "Graze Factor"
      Height          =   252
      Index           =   20
      Left            =   3120
      TabIndex        =   42
      Top             =   4200
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Y Coordinate"
      Height          =   252
      Index           =   19
      Left            =   3120
      TabIndex        =   39
      Top             =   3840
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Y Goal"
      Height          =   252
      Index           =   18
      Left            =   3120
      TabIndex        =   37
      Top             =   3480
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Y speed"
      Height          =   252
      Index           =   17
      Left            =   3120
      TabIndex        =   35
      Top             =   3120
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Max Speed"
      Height          =   252
      Index           =   16
      Left            =   3120
      TabIndex        =   33
      Top             =   2760
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Rest"
      Height          =   252
      Index           =   15
      Left            =   3120
      TabIndex        =   31
      Top             =   2400
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Injury"
      Height          =   252
      Index           =   14
      Left            =   3120
      TabIndex        =   29
      Top             =   2040
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Size to breed"
      Height          =   252
      Index           =   13
      Left            =   3120
      TabIndex        =   27
      Top             =   1680
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Maximum Food"
      Height          =   252
      Index           =   12
      Left            =   3120
      TabIndex        =   25
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Maximum Size"
      Height          =   252
      Index           =   11
      Left            =   3120
      TabIndex        =   23
      Top             =   960
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Maximum Age"
      Height          =   252
      Index           =   10
      Left            =   3120
      TabIndex        =   21
      Top             =   600
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "X Coordinate"
      Height          =   252
      Index           =   9
      Left            =   240
      TabIndex        =   19
      Top             =   3840
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "X Goal"
      Height          =   252
      Index           =   8
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "X speed"
      Height          =   252
      Index           =   7
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Species"
      Height          =   252
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Top             =   2760
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Metabolism"
      Height          =   252
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Health"
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Eat Factor"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Food"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Size"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Age"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1452
   End
End
Attribute VB_Name = "CreatureDataForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseBtn_Click()

Me.Hide
End Sub

