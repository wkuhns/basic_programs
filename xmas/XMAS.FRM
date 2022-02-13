VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   2070
   ClientTop       =   1635
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7515
   ScaleWidth      =   5805
   Begin VB.CommandButton Command2 
      Caption         =   "Save It"
      Height          =   372
      Left            =   3000
      TabIndex        =   21
      Top             =   1080
      Width           =   972
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do It"
      Height          =   372
      Left            =   3000
      TabIndex        =   20
      Top             =   600
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   19
      Left            =   1800
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3840
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   18
      Left            =   1800
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3480
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   17
      Left            =   1800
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3120
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   16
      Left            =   1800
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2760
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   15
      Left            =   1800
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2400
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   14
      Left            =   1800
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2040
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   13
      Left            =   1800
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1680
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   12
      Left            =   1800
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   11
      Left            =   1800
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   960
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   10
      Left            =   1800
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   600
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   9
      Left            =   600
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3840
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   8
      Left            =   600
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3480
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   7
      Left            =   600
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3120
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   6
      Left            =   600
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2760
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   5
      Left            =   600
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2400
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   4
      Left            =   600
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2040
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   3
      Left            =   600
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1680
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1320
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   972
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim names(20) As String
Dim bad(5, 20) As Integer
Dim used(20) As Integer
Dim picks(20) As Integer

Private Sub Command1_Click()

    Dim i As Integer
    Dim pick As Integer
    Dim tries As Integer
    
    Randomize

restart:
    tries = 0
    For i = 1 To 8
        used(i) = 0
    Next i
    
    'For i = 1 To 10
    ' Skip Mei-li this year
    For i = 1 To 8
        
        'pick = Int(Rnd() * 10) + 1
        pick = Int(Rnd() * 8) + 1
        
        While (tries < 100) And (used(pick) = 1 Or pick = i Or pick = bad(1, i) Or pick = bad(2, i) Or pick = bad(3, i))
            'pick = Int(Rnd() * 10) + 1
            pick = Int(Rnd() * 8) + 1
            tries = tries + 1
        Wend
        If tries >= 100 Then
 '           Stop
            GoTo restart
        Else
           picks(i) = pick
           used(pick) = 1
        End If
nexti:
    Next i
End Sub

Private Sub Command2_Click()
            
Dim i As Integer

 Open "history.dat" For Output As #1

For i = 1 To 8
        
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print "      Merry Christmas, " + names(i) + "! You have drawn " + names(picks(i)) + " for Christmas."
   Printer.NewPage
   Print #1, picks(i), bad(2, i)
'    Debug.Print "      Merry Christmas, " + names(i) + "! You have drawn " + names(picks(i)) + " for Christmas."
nexti:
Next i

Printer.EndDoc
Close #1

End Sub


Private Sub Form_Load()

    Dim i As Integer
    
    names(1) = "Mom"
    names(2) = "Bill"
    names(3) = "Cindy"
    names(4) = "Buzz"
    names(5) = "Carroll"
    names(6) = "Don"
    names(7) = "Hal"
    names(8) = "Bre"
    
    ' don't allow spouses to get each other
    bad(1, 1) = 0
    bad(1, 2) = 3
    bad(1, 3) = 2
    bad(1, 4) = 5
    bad(1, 5) = 4
    bad(1, 6) = 0
    bad(1, 7) = 8
    bad(1, 8) = 7
    
    Open "history.dat" For Input As #1
    For i = 1 To 8
        Input #1, bad(2, i), bad(3, i)
        Text1(i - 1).Text = names(i)
        Text1(i + 9).Text = ""
    Next i
    Close #1
    
    
End Sub


