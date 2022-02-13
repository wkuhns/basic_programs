VERSION 5.00
Begin VB.Form BuyForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Purchase Order"
   ClientHeight    =   2832
   ClientLeft      =   5916
   ClientTop       =   1800
   ClientWidth     =   6420
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
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2832
   ScaleWidth      =   6420
   Begin VB.TextBox TallyBox 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   5880
      TabIndex        =   20
      Top             =   240
      Width           =   492
   End
   Begin VB.CommandButton DoneButton 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   372
      Left            =   5520
      TabIndex        =   19
      Top             =   2400
      Width           =   852
   End
   Begin VB.TextBox BalanceBox 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   4200
      TabIndex        =   17
      Top             =   240
      Width           =   1572
   End
   Begin VB.TextBox TotalBox 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   2640
      TabIndex        =   6
      Top             =   240
      Width           =   972
   End
   Begin VB.CommandButton OKButton 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   372
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   372
   End
   Begin VB.TextBox QtyBox 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   612
   End
   Begin VB.ComboBox UnitList 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1812
   End
   Begin VB.CommandButton DeleteButton 
      Appearance      =   0  'Flat
      Caption         =   "Delete"
      Height          =   372
      Left            =   3960
      TabIndex        =   1
      Top             =   2400
      Width           =   852
   End
   Begin VB.ListBox BuyList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   3960
      TabIndex        =   0
      Top             =   600
      Width           =   2412
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Balance"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   4680
      TabIndex        =   18
      Top             =   0
      Width           =   972
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   7
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   3492
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   3492
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   3492
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Width           =   3492
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   3492
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   3492
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   3492
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   3492
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2760
      TabIndex        =   8
      Top             =   0
      Width           =   732
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Quantity"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1920
      TabIndex        =   7
      Top             =   0
      Width           =   732
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Item"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   972
   End
End
Attribute VB_Name = "BuyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim balance As Integer
Dim tally As Integer

Private Sub DeleteButton_Click()

'   The user wants to delete the selected item from the order list

    balance = balance + specs(Val(Left$(BuyList.Text, 2))).buycost * Val(Mid$(BuyList.Text, 4, 2))
    tally = tally + Val(Mid$(BuyList.Text, 4, 2))
    BuyList.RemoveItem BuyList.ListIndex
    BalanceBox.Text = Str$(balance)
    TallyBox.Text = Str$(tally)

End Sub

Private Sub DoneButton_Click()

'   User is done buying things.

    Dim i, j, k, l As Integer
    Dim unit As unitstruct

    k = 1
    ' For each line on the order list, buy the specified quantity
    For i = 0 To BuyList.ListCount - 1
        For j = 0 To Val(Mid$(BuyList.List(i), 4, 2)) - 1
            unit.type = Val(Left$(BuyList.List(i), 2))
            unit.side = side
            unit.Index = k
            unit.health = 100
            If side = us Then
                unit.x = a1base.x
                unit.y = a1base.y
                unit.dx = a1base.x
                unit.dy = a1base.y
                unit.ocount = 0
                army(0, k) = unit
                specs(unit.type).count = specs(unit.type).count + 1
                SendUnitInfo army(0, k)
            Else
                unit.x = a2base.x
                unit.y = a2base.y
                unit.dx = a2base.x
                unit.dy = a2base.y
                unit.ocount = 0
                army(1, k) = unit
            End If
            k = k + 1
        Next j
    Next i
    UpdateDispBoxes
    BuyList.Clear
    Unload BuyForm

End Sub

Private Sub Form_Load()

    Dim i As Integer

    UnitList.Clear

    ' Fill unit list. Unit 0 is the general, so don't let him buy another
    For i = 1 To UBound(specs) - 1
        UnitList.AddItem specs(i).name
    Next i

    UnitList.ListIndex = 1                  ' set initial unit
    balance = 10000
    tally = 49
    BalanceBox.Text = Str$(balance)
    TallyBox.Text = Str$(tally)

End Sub

Private Sub OKButton_Click()

'   User is happy with the unit specified and the quantity specified.
'   Transfer it to the BuyList

    Dim buff As String
    Dim qty, cost As Integer

    qty = Val(QtyBox.Text)

    If qty > tally Then
        qty = tally
    End If

    cost = specs(UnitList.ListIndex + 1).buycost

    If cost * qty > balance Then
        qty = Int(balance / cost)
    End If

    QtyBox.Text = Str$(qty)
    TotalBox.Text = Str$(cost * qty)

    tally = tally - qty
    balance = balance - cost * qty

    buff = Format$(UnitList.ListIndex + 1, "00 ")
    buff = buff + Format$(qty, "00 ")
    buff = buff + Format$(specs(UnitList.ListIndex + 1).name, "!@@@@@@@@ ")
    buff = buff + Str$(cost * qty)
    BuyList.AddItem buff

    BalanceBox.Text = Str$(balance)
    TallyBox.Text = Str$(tally)

End Sub

Private Sub QtyBox_Change()

    TotalBox.Text = Str$(Val(QtyBox.Text) * specs(UnitList.ListIndex + 1).buycost)

End Sub

Private Sub UnitList_Click()

    Dim i As Integer

    i = UnitList.ListIndex + 1
    SpecLabel(0).Caption = "Cost:       " + Str$(specs(i).buycost)
    SpecLabel(1).Caption = "W strength: " + Str$(specs(i).wstrength)
    SpecLabel(2).Caption = "W range:    " + Str$(specs(i).wrange)
    SpecLabel(3).Caption = "Armor:      " + Str$(specs(i).armor)
    SpecLabel(4).Caption = "Speed:      " + Str$(specs(i).speed)
    SpecLabel(5).Caption = "Range:      " + Str$(specs(i).range)
    SpecLabel(6).Caption = "Vision:     " + Str$(specs(i).vision)
    SpecLabel(7).Caption = "Cost:       " + Str$(specs(i).buycost)

    QtyBox.Text = "1"

End Sub

