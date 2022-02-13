VERSION 5.00
Begin VB.Form SubMap 
   Caption         =   "Form1"
   ClientHeight    =   2880
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   3840
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Enabled         =   0   'False
      Height          =   132
      Left            =   2520
      ScaleHeight     =   84
      ScaleWidth      =   84
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2640
      LargeChange     =   20
      Left            =   3600
      Max             =   80
      TabIndex        =   1
      Top             =   0
      Width           =   240
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   252
      LargeChange     =   20
      Left            =   0
      Max             =   80
      TabIndex        =   0
      Top             =   2640
      Width           =   3600
   End
End
Attribute VB_Name = "SubMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'top left x,y coords in cells
    Public tx As Integer
    Public ty As Integer
    'bottom right x,y coords
    Public bx As Integer
    Public by As Integer
    'height and width of selection
    Public dx As Integer
    Public dy As Integer
    'text box offsets
    Public xboxoffset As Single
    Public yboxoffset As Single
    Public textboxheight As Single
    Public textboxwidth As Single

' This call creates us. We initialize globals here
Public Sub SetSubMap(mode As Integer, x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)

    Dim deficit As Integer
    'Sort out top left, bottom right
    If x1 < x2 Then
        tx = x1
        bx = x2
    Else
        tx = x2
        bx = x1
    End If
    
    If y1 < y2 Then
        ty = y1
        by = y2
    Else
        ty = y2
        by = y1
    End If
    
    dx = bx - tx
    dy = by - ty
    
    Me.Show
    ' Correct height and width to 20 pixels per cell
    deficit = ((dx + 1) * 20) - Me.ScaleWidth
    If deficit <> 0 Then
        Me.Width = Me.Width + deficit * 12 + VScroll1.Width
    End If
    deficit = ((dy + 1) * 20) - Me.ScaleHeight
    If deficit <> 0 Then
        Me.Height = Me.Height + deficit * 12 + HScroll1.Height
    End If
    
    ' Set up scroll bars
    Me.ScaleHeight = 20 * (dy + 1) + HScroll1.Height
    Me.ScaleWidth = 20 * (dx + 1) + VScroll1.Width
    HScroll1.min = 0
    VScroll1.min = 0
    HScroll1.max = axis - (dx)
    VScroll1.max = axis - (dy)
    
    HScroll1.Value = tx
    VScroll1.Value = ty
    HScroll1.LargeChange = dx
    VScroll1.LargeChange = dy
    
    ' Text box metrics
    textboxwidth = Me.TextWidth("D") + 1
    textboxheight = Me.TextHeight("D") - 2
    xboxoffset = textboxwidth / -2
    yboxoffset = textboxheight / -2
    
    DrawMap mode
    ' Show user what area detail map shows
    Call MapForm.DrawSelectionBox(tx, ty, bx, by)

End Sub

' mode tells us whether we're displaying altitude, difficulty, or cover
Private Sub DrawMap(mode As Integer)
    
    Dim X As Integer
    Dim Y As Integer
    Dim h As Integer
    Dim xp As Integer
    
    Me.AutoRedraw = True
    For X = tx To bx
        xp = (X - tx) * 20
        For Y = ty To by
            h = Int(terrain(X, Y).a / 10)     ' quantize
            If h > 0 And rain(X, Y) > 10 Then h = 0
            Me.Line (xp, (Y - ty) * 20)-Step(20, 20), mc(h), BF
        Next Y
    Next X

    DisplayRivers
    Me.AutoRedraw = False
    
End Sub
' We display river cells with 'x' through the whole cell. The width of the lines
' indicates the width of the river.
Sub DisplayRivers()
    Dim X As Integer
    Dim Y As Integer
    Dim yp As Integer
    Dim h As Integer
    Dim dkh As Integer
    Dim kw As Integer

    'MapForm.DrawWidth = 2
    ' Draw colors
    For Y = 0 To dy
        yp = Y * 20
        For X = 0 To dx
            ' skip if no river
            If terrain(X + tx, Y + ty).d <= 10 Then
                GoTo nextx
            End If
            Me.DrawWidth = terrain(X + tx, Y + ty).d - 10
            If terrain(X + tx, Y + ty).d > 10 Then
                Me.Line (X * 20, yp)-Step(20, 20), mc(0)
                Me.Line ((X + 1) * 20, yp)-Step(-20, 20), mc(0)
            End If
nextx:
        Next X
    Next Y
    
End Sub
Sub DisplayUnits()
    
    Dim u As Integer
    Dim ux As Single
    Dim uy As Single
    
    Me.AutoRedraw = False
    Me.Refresh
    For u = 0 To asize
        If army(us, u).health <= 0 Then
            GoTo nextu
        End If
        If army(us, u).X < tx Or army(us, u).X > bx Then
            GoTo nextu
        End If
        If army(us, u).Y < ty Or army(us, u).Y > by Then
            GoTo nextu
        End If
        ux = (army(us, u).X - tx) * 20
        uy = (army(us, u).Y - ty) * 20
        Me.FillColor = WHITE
        Me.Line (ux + xboxoffset, uy + yboxoffset)-Step(textboxwidth, textboxheight), WHITE, B
        Me.CurrentX = ux - 5
        Me.CurrentY = uy - 8
        Me.Print Left$(specs(army(us, u).type).name, 1)
nextu:
    Next u

End Sub
Private Sub Form_Load()

HScroll1.min = 0
HScroll1.max = axis - 21
VScroll1.min = 0
VScroll1.max = axis - 21

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
DisplayUnits
End Sub

Private Sub Form_Resize()


HScroll1.Width = Me.ScaleWidth - VScroll1.Width
HScroll1.Top = Me.ScaleHeight - HScroll1.Height
VScroll1.Height = Me.ScaleHeight - HScroll1.Height
VScroll1.Left = Me.ScaleWidth - VScroll1.Width


End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call MapForm.DrawSelectionBox(tx, ty, bx, by)

End Sub

Private Sub HScroll1_Change()

    Call MapForm.DrawSelectionBox(tx, ty, bx, by)
    tx = HScroll1.Value
    bx = HScroll1.Value + dx
    DrawMap (0)
    Call MapForm.DrawSelectionBox(tx, ty, bx, by)
    'Picture1.SetFocus

End Sub

Private Sub VScroll1_Change()

    Call MapForm.DrawSelectionBox(tx, ty, bx, by)
    ty = VScroll1.Value
    by = VScroll1.Value + dy
    DrawMap (0)
    Call MapForm.DrawSelectionBox(tx, ty, bx, by)
    'Picture1.SetFocus
End Sub
