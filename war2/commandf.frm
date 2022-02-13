VERSION 5.00
Begin VB.Form CommandForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Unit Status"
   ClientHeight    =   3444
   ClientLeft      =   4560
   ClientTop       =   468
   ClientWidth     =   6120
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
   ScaleHeight     =   3444
   ScaleWidth      =   6120
   Begin VB.CommandButton OrderButton 
      Appearance      =   0  'Flat
      Caption         =   "Orders"
      Height          =   372
      Left            =   1680
      TabIndex        =   0
      Top             =   3000
      Width           =   852
   End
   Begin VB.ListBox UnitBox 
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
      Height          =   1944
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   16
      Top             =   0
      Width           =   3012
   End
   Begin VB.CommandButton HomeButton 
      Appearance      =   0  'Flat
      Caption         =   "Home"
      Height          =   372
      Left            =   2640
      TabIndex        =   24
      Top             =   3000
      Width           =   732
   End
   Begin VB.CommandButton RevokeButton 
      Appearance      =   0  'Flat
      Caption         =   "Revoke Orders"
      Height          =   372
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   1452
   End
   Begin VB.CommandButton EndButton 
      Appearance      =   0  'Flat
      Caption         =   "End"
      Height          =   372
      Left            =   4320
      TabIndex        =   22
      Top             =   3000
      Width           =   732
   End
   Begin VB.HScrollBar FollowBar 
      Height          =   252
      Left            =   960
      Max             =   100
      TabIndex        =   21
      Top             =   2640
      Width           =   2052
   End
   Begin VB.HScrollBar RetreatBar 
      Height          =   252
      Left            =   960
      Max             =   100
      TabIndex        =   20
      Top             =   2400
      Width           =   2052
   End
   Begin VB.HScrollBar AttackBar 
      Height          =   252
      Left            =   960
      Max             =   100
      TabIndex        =   19
      Top             =   2160
      Width           =   2052
   End
   Begin VB.CommandButton PathButton 
      Appearance      =   0  'Flat
      Caption         =   "Path"
      Height          =   372
      Left            =   3480
      TabIndex        =   18
      Top             =   3000
      Width           =   732
   End
   Begin VB.CommandButton DoneButton 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   372
      Left            =   5160
      TabIndex        =   14
      Top             =   3000
      Width           =   732
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Follow"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   852
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Retreat"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   852
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Attack"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   852
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   11
      Left            =   3120
      TabIndex        =   17
      Top             =   0
      Width           =   3012
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   10
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   3012
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   9
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Width           =   3012
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   8
      Left            =   3120
      TabIndex        =   10
      Top             =   720
      Width           =   3012
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   7
      Left            =   3120
      TabIndex        =   15
      Top             =   2760
      Width           =   3012
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   6
      Left            =   3120
      TabIndex        =   8
      Top             =   2520
      Width           =   3012
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   5
      Left            =   3120
      TabIndex        =   7
      Top             =   2280
      Width           =   3012
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   4
      Left            =   3120
      TabIndex        =   6
      Top             =   2040
      Width           =   3012
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   3
      Left            =   3120
      TabIndex        =   5
      Top             =   1800
      Width           =   3012
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   2
      Left            =   3120
      TabIndex        =   4
      Top             =   1560
      Width           =   3012
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   1320
      Width           =   3012
   End
   Begin VB.Label SpecLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   3012
   End
End
Attribute VB_Name = "CommandForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AttackBar_Change()
    
    Dim i As Integer

    For i = 0 To CommandForm!UnitBox.ListCount - 1
        If CommandForm!UnitBox.Selected(i) Then
            If side = us Then
                army(0, Val(Left$(UnitBox.List(i), 2))).attack = AttackBar.Value
            Else
                army(1, Val(Left$(UnitBox.List(i), 2))).attack = AttackBar.Value
            End If
        End If
    Next i


End Sub

Private Sub DoneButton_Click()

    UnitBox.Clear
    CommandForm.Hide
    UnitWaiting = False
    MapForm.MousePointer = 0

End Sub

Private Sub EndButton_Click()

    UnitWaiting = False
    UnitBox.Enabled = True
    MapForm.MousePointer = 0

End Sub

Private Sub FollowBar_Change()
    
    Dim i As Integer

    For i = 0 To CommandForm!UnitBox.ListCount - 1
        If CommandForm!UnitBox.Selected(i) Then
            If side = us Then
                army(0, Val(Left$(UnitBox.List(i), 2))).follow = FollowBar.Value
            Else
                army(1, Val(Left$(UnitBox.List(i), 2))).follow = FollowBar.Value
            End If
        End If
    Next i

End Sub

Private Sub HomeButton_Click()
    
    Dim i As Integer

    For i = 0 To CommandForm!UnitBox.ListCount - 1
        If CommandForm!UnitBox.Selected(i) Then
            If side = us Then
                InsertOrder army(0, Val(Left$(UnitBox.List(i), 2))), "M", a1base.x, a1base.y
            Else
                InsertOrder army(1, Val(Left$(UnitBox.List(i), 2))), "M", a2base.x, a2base.y
            End If
        End If
    Next i


End Sub

Private Sub OrderButton_Click()

    Dim unit As unitstruct
    Dim i As Integer
    Dim selcount As Integer

    OrderForm.Show
    OrderForm!OrderList.Clear

    If side = us Then
        unit = army(0, Val(Left$(UnitBox.Text, 2)))
    Else
        unit = army(1, Val(Left$(UnitBox.Text, 2)))
    End If
    
    selcount = 0
    For i = 0 To OrderForm!OrderList.ListCount
        If UnitBox.Selected(i) Then
            selcount = selcount + 1
        End If
    Next i

    ' if just one unit selected, show his orders.

    If selcount = 1 Then
        For i = 0 To unit.ocount - 1
            OrderForm!OrderList.AddItem unit.orders(i).command + Format$(unit.orders(i).n1, " 000.00") + Format$(unit.orders(i).n2, " 000.00")
        Next i
    End If

End Sub

Private Sub PaintCommandForm()
    
    Dim i As Integer
    Dim unit As unitstruct
    Dim scount As Integer

'    UnitBox.Height = SpecLabel(11).Top
    
    scount = 0
    For i = 0 To UnitBox.ListCount - 1
        If UnitBox.Selected(i) = True Then
            scount = scount + 1
        End If
    Next i

    If scount > 1 Then

        For i = 0 To 11
            SpecLabel(i).Visible = False
        Next i
'        Label1.Visible = False
        Label2.Enabled = False
        Label3.Enabled = False
        Label4.Enabled = False

    Else
        
        For i = 0 To 11
            SpecLabel(i).Visible = True
        Next i
'        Label1.Visible = True
        Label2.Enabled = True
        Label3.Enabled = True
        Label4.Enabled = True
        i = Val(Left$(UnitBox.Text, 2))
        If side = us Then
            unit = army(0, i)
        Else
            unit = army(1, i)
        End If

'        CommandForm!Label1.Caption = "Unit: " + specs(Unit.type).name
        CommandForm!SpecLabel(0).Caption = "Max Speed      : " + Str$(specs(unit.type).speed)
        CommandForm!SpecLabel(1).Caption = "Armor          : " + Str$(specs(unit.type).armor)
        CommandForm!SpecLabel(2).Caption = "Weapon Range   : " + Str$(specs(unit.type).wrange)
        CommandForm!SpecLabel(3).Caption = "Weapon Strength: " + Str$(specs(unit.type).wstrength)
        CommandForm!SpecLabel(4).Caption = "Accuracy       : " + Str$(specs(unit.type).accuracy)
        CommandForm!SpecLabel(5).Caption = "Unit Range     : " + Str$(specs(unit.type).range)
        CommandForm!SpecLabel(6).Caption = "Fuel Capacity  : " + Str$(specs(unit.type).fuelcap)
        CommandForm!SpecLabel(7).Caption = "Vision         : " + Str$(specs(unit.type).vision)
    
        CommandForm!SpecLabel(8).Caption = "Health      : " + Str$(unit.health)
        CommandForm!SpecLabel(9).Caption = "Speed       : " + Str$(unit.speed)
        CommandForm!SpecLabel(10).Caption = "Fuel        : " + Str$(unit.fuel)
'        CommandForm!SpecLabel(11).Caption = "Position    : " + Format$(Unit.x, "##.#") + Format$(Unit.y, ", ##.#")
        CommandForm!SpecLabel(11).Caption = "Destination : " + Format$(unit.dx, "##.#") + Format$(unit.dy, ", ##.#")
        CommandForm!AttackBar.Value = unit.attack
        CommandForm!RetreatBar.Value = unit.Retreat
        CommandForm!FollowBar.Value = unit.follow
    End If
End Sub

Private Sub PathButton_Click()
    
    UnitWaiting = True
    UnitBox.Enabled = False
    MapForm.MousePointer = 2


End Sub

Private Sub RetreatBar_Change()
    
    Dim i As Integer

    For i = 0 To CommandForm!UnitBox.ListCount - 1
        If CommandForm!UnitBox.Selected(i) Then
            If side = us Then
                army(0, Val(Left$(UnitBox.List(i), 2))).Retreat = RetreatBar.Value
            Else
                army(1, Val(Left$(UnitBox.List(i), 2))).Retreat = RetreatBar.Value
            End If
        End If
    Next i

End Sub

Private Sub RevokeButton_Click()

    Dim i As Integer

    For i = 0 To CommandForm!UnitBox.ListCount - 1
        If CommandForm!UnitBox.Selected(i) Then
            If side = us Then
                army(0, Val(Left$(UnitBox.List(i), 2))).ocount = 0
                army(0, Val(Left$(UnitBox.List(i), 2))).dx = army(0, Val(Left$(UnitBox.List(i), 2))).x
                army(0, Val(Left$(UnitBox.List(i), 2))).dy = army(0, Val(Left$(UnitBox.List(i), 2))).y
            Else
                army(1, Val(Left$(UnitBox.List(i), 2))).ocount = 0
                army(1, Val(Left$(UnitBox.List(i), 2))).dx = army(1, Val(Left$(UnitBox.List(i), 2))).x
                army(1, Val(Left$(UnitBox.List(i), 2))).dy = army(1, Val(Left$(UnitBox.List(i), 2))).y
            End If
        End If
    Next i

    If side = us Then
    End If

End Sub

Private Sub UnitBox_Click()

    PaintCommandForm

End Sub

Private Sub UnitBox_DblClick()

    Dim buff As String

    buff = UnitBox.Text
    UnitBox.Clear
    UnitBox.AddItem buff
    UnitBox.Selected(0) = True
    PaintCommandForm

End Sub

