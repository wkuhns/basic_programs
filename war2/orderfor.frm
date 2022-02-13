VERSION 5.00
Begin VB.Form OrderForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Orders"
   ClientHeight    =   3900
   ClientLeft      =   8820
   ClientTop       =   4188
   ClientWidth     =   3132
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
   ScaleHeight     =   3900
   ScaleWidth      =   3132
   Begin VB.CommandButton PathButton 
      Appearance      =   0  'Flat
      Caption         =   "Path"
      Height          =   372
      Left            =   2280
      TabIndex        =   13
      Top             =   3000
      Width           =   732
   End
   Begin VB.HScrollBar AttackBar 
      Height          =   252
      Left            =   960
      Max             =   100
      TabIndex        =   9
      Top             =   1920
      Width           =   2172
   End
   Begin VB.HScrollBar RetreatBar 
      Height          =   252
      Left            =   960
      Max             =   100
      TabIndex        =   8
      Top             =   2280
      Width           =   2172
   End
   Begin VB.HScrollBar FollowBar 
      Height          =   252
      Left            =   960
      Max             =   100
      TabIndex        =   7
      Top             =   2640
      Width           =   2172
   End
   Begin VB.CommandButton CancelButton 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   372
      Left            =   1560
      TabIndex        =   6
      Top             =   3480
      Width           =   732
   End
   Begin VB.CommandButton DelButton 
      Appearance      =   0  'Flat
      Caption         =   "Delete"
      Height          =   372
      Left            =   840
      TabIndex        =   5
      Top             =   3480
      Width           =   612
   End
   Begin VB.CommandButton AddButton 
      Appearance      =   0  'Flat
      Caption         =   "Add"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   612
   End
   Begin VB.CommandButton DoneButton 
      Appearance      =   0  'Flat
      Caption         =   "Done"
      Height          =   372
      Left            =   2400
      TabIndex        =   3
      Top             =   3480
      Width           =   612
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1560
      TabIndex        =   2
      Top             =   3000
      Width           =   492
   End
   Begin VB.ComboBox CMDBox 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3000
      Width           =   1332
   End
   Begin VB.ListBox OrderList 
      Appearance      =   0  'Flat
      Height          =   1560
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   0
      Width           =   3132
   End
   Begin VB.Label Labelx 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Follow"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   852
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Retreat"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   852
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Attack"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   852
   End
End
Attribute VB_Name = "OrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddButton_Click()

        OrderList.AddItem Left$(CMDBox.Text, 1) + Format$(Val(Text1.Text), " 000.00") + " 000.00"

End Sub

Private Sub AttackBar_Change()

    SendOrders

End Sub

Private Sub CancelButton_Click()

        OrderForm.Hide

End Sub

Private Sub DelButton_Click()
        
    Dim i As Integer

    For i = (OrderList.ListCount - 1) To 0 Step -1
        If OrderList.Selected(i) Then
            OrderList.RemoveItem i
        End If
    Next i

End Sub

Private Sub DoneButton_Click()

        Dim i, j, u As Integer
        Dim cmd As String
        Dim n1, n2 As Single

        For i = 0 To CommandForm!UnitBox.ListCount - 1
                If CommandForm!UnitBox.Selected(i) Then
                        u = Val(Left$(CommandForm!UnitBox.List(i), 2))
                        If side = us Then
                                army(0, u).ocount = 0
                        Else
                                army(1, u).ocount = 0
                        End If

                        For j = 0 To OrderList.ListCount - 1
                                cmd = Left$(OrderList.List(j), 1)
                                n1 = CDbl(Mid$(OrderList.List(j), 3, 6))
                                n2 = CDbl(Mid$(OrderList.List(j), 10, 6))
                                If side = us Then
                                        AddOrder army(0, u), cmd, CSng(n1), CSng(n2)
                                Else
                                        AddOrder army(1, u), cmd, CSng(n1), CSng(n2)
                                End If
                                WriteCCC "Order: " + Str$(u) + " " + cmd + Str$(n1) + ", " + Str$(n2)
                        Next j
                End If
        Next i

        OrderForm.Hide

End Sub

Private Sub FollowBar_Change()
    
    SendOrders

End Sub

Private Sub Form_Load()


        CMDBox.Clear
        CMDBox.AddItem "Move"
        CMDBox.AddItem "Attack"
        CMDBox.AddItem "Follow"
        CMDBox.AddItem "Wait"
        CMDBox.AddItem "Retreat"
        CMDBox.AddItem "Send"
        CMDBox.AddItem "Camoflage"

        CMDBox.ListIndex = 1

End Sub

Private Sub OrderList_Click()

        Text1.Text = Mid$(OrderList.Text, 3, 6)

        Select Case Left$(OrderList.Text, 1)
                Case "M"
                        CMDBox.ListIndex = 0
                Case "X"
                        CMDBox.ListIndex = 0
                Case "A"
                        CMDBox.ListIndex = 1
                Case "F"
                        CMDBox.ListIndex = 2
                Case "W"
                        CMDBox.ListIndex = 3
                Case "R"
                        CMDBox.ListIndex = 4
                Case "S"
                        CMDBox.ListIndex = 5
                Case "C"
                        CMDBox.ListIndex = 6
        End Select
End Sub

Private Sub PathButton_Click()

    UnitWaiting = True
    CommandForm!UnitBox.Enabled = False
    MapForm.MousePointer = 2

End Sub

Private Sub RetreatBar_Change()
    
    SendOrders

End Sub

Private Sub SendOrders()
        
        Dim i, j, u As Integer
        Dim cmd As String
        Dim n1, n2 As Single

        For i = 0 To CommandForm!UnitBox.ListCount - 1
            If CommandForm!UnitBox.Selected(i) Then
                u = Val(Left$(CommandForm!UnitBox.List(i), 2))

                cmd = "A"
                n1 = AttackBar.Value
                n2 = 0
                If side = us Then
                    army(0, u).attack = AttackBar.Value
                    army(0, u).follow = FollowBar.Value
                    army(0, u).Retreat = RetreatBar.Value
                Else
                    army(1, u).attack = AttackBar.Value
                    army(1, u).follow = FollowBar.Value
                    army(1, u).Retreat = RetreatBar.Value
                End If
            End If
        Next i

End Sub

