VERSION 2.00
Begin Form DetailForm 
   Caption         =   "Detail"
   ClientHeight    =   3960
   ClientLeft      =   6840
   ClientTop       =   564
   ClientWidth     =   3816
   Height          =   4380
   Left            =   6792
   LinkTopic       =   "Form2"
   ScaleHeight     =   3960
   ScaleWidth      =   3816
   Top             =   192
   Width           =   3912
   Begin CommandButton Command1 
      Caption         =   "Done"
      Height          =   372
      Left            =   2880
      TabIndex        =   3
      Top             =   3600
      Width           =   972
   End
   Begin TextBox ylabel 
      Height          =   372
      Left            =   960
      TabIndex        =   2
      Top             =   3600
      Width           =   972
   End
   Begin TextBox xlabel 
      Height          =   372
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Width           =   972
   End
   Begin PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3612
      Left            =   0
      ScaleHeight     =   3588
      ScaleWidth      =   3828
      TabIndex        =   0
      Top             =   0
      Width           =   3852
   End
End
Option Explicit

Dim xbase, ybase As Integer         ' x,y of top left cell

Sub Command1_Click ()
    DetailForm.Hide
End Sub

Sub DrawDetail (x, y As Integer)
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim H As Integer
    Dim dh As Integer
    Dim dw As Integer

    dh = DetailForm!Picture1.Height / 10
    dw = DetailForm!Picture1.Width / 10

    x = Int(x / mapdx)
    y = Int(y / mapdy)

    x = x - 5
    y = y - 5
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    If x > (axis - 10) Then x = axis - 10
    If y > (axis - 10) Then y = axis - 10
    
    ' draw selection box on MapForm
    
'    MapForm!Picture1.AutoRedraw = False
'    MapForm!Picture1.Cls
'    MapForm!Picture1.Line (x * mapdx, y * dh1)-Step(MapForm!Picture1.Width / 10, MapForm!Picture1.Height / 10), RGB(0, 0, 0), B
'    MapForm!Picture1.AutoRedraw = True

    DetailForm.Show
    
    DetailForm!xlabel.Text = Str$(x)
    DetailForm!ylabel.Text = Str$(y)

    For i = 0 To 9
        For j = 0 To 9
            If x + i < axis And y + j < axis Then
                H = Int(t(x + i, y + j) / 10)' quantize
                DetailForm!Picture1.Line (i * dw, j * dh)-Step(dw, dh), mc(H), BF
                For k = 0 To asize - 1
                    If side = us Then
                        If a1(k).health > 0 And Int(a1(k).x) = x + i And Int(a1(k).y) = y + j Then
                            DetailForm!Picture1.Line (i * dw, j * dh)-Step(dw, dh), 0
                            DetailForm!Picture1.Line (i * dw + dw, j * dh)-Step(-dw, dh), 0
                            Exit For
                        End If
                    Else
                        If a2(k).health > 0 And Int(a2(k).x) = x + i And Int(a2(k).y) = y + j Then
                            DetailForm!Picture1.Line (i * dw, j * dh)-Step(dw, dh), 0
                            DetailForm!Picture1.Line (i * dw + dw, j * dh)-Step(-dw, dh), 0
                            Exit For
                        End If
                    End If
                Next k
            End If
        Next j
    Next i
    
End Sub

Sub Picture1_MouseDown (Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim dw, dh, i, j As Integer

    dw = Picture1.Width / 10
    dh = Picture1.Height / 10

    PickForm.Show
    PickForm!PickList.Clear

    x = Int(x / dw) + Val(xlabel.Text)
    y = Int(y / dh) + Val(ylabel.Text)

    i = 0
    j = 0
    For i = 0 To asize - 1
        If side = us Then
            If Int(a1(i).x) = x And Int(a1(i).y) = y Then
                Call AddListItem(a1(i))
            End If
        Else
            If Int(a2(i).x) = x And Int(a2(i).y) = y Then
                Call AddListItem(a2(i))
            End If
        End If
    Next i

End Sub

