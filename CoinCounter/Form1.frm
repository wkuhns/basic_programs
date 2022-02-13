VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2244
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   2304
   LinkTopic       =   "Form1"
   ScaleHeight     =   2244
   ScaleWidth      =   2304
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Coins 
      Height          =   264
      Index           =   3
      Left            =   800
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "0"
      Top             =   1727
      Width           =   616
   End
   Begin VB.TextBox Coins 
      Height          =   264
      Index           =   2
      Left            =   800
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "0"
      Top             =   1452
      Width           =   616
   End
   Begin VB.TextBox Coins 
      Height          =   264
      Index           =   1
      Left            =   800
      MaxLength       =   10
      TabIndex        =   8
      Text            =   "0"
      Top             =   1177
      Width           =   616
   End
   Begin VB.TextBox Coins 
      Height          =   288
      Index           =   0
      Left            =   800
      MaxLength       =   25
      TabIndex        =   2
      Text            =   "0"
      Top             =   902
      Width           =   616
   End
   Begin VB.TextBox Balance 
      Height          =   264
      Left            =   0
      LinkTimeout     =   0
      TabIndex        =   1
      Text            =   "86"
      Top             =   480
      Width           =   616
   End
   Begin VB.Label Label1 
      Caption         =   "Enter amount in 'Balance' box, then press 'tab' key"
      Height          =   408
      Left            =   -12
      TabIndex        =   10
      Top             =   60
      Width           =   2268
   End
   Begin VB.Label Coin 
      Alignment       =   1  'Right Justify
      Caption         =   "Balance"
      Height          =   253
      Index           =   4
      Left            =   671
      TabIndex        =   9
      Top             =   550
      Width           =   616
   End
   Begin VB.Label Coin 
      Alignment       =   1  'Right Justify
      Caption         =   "Pennies"
      Height          =   253
      Index           =   3
      Left            =   1
      TabIndex        =   7
      Top             =   1705
      Width           =   616
   End
   Begin VB.Label Coin 
      Alignment       =   1  'Right Justify
      Caption         =   "Nickels"
      Height          =   253
      Index           =   2
      Left            =   5
      TabIndex        =   4
      Top             =   1441
      Width           =   616
   End
   Begin VB.Label Coin 
      Alignment       =   1  'Right Justify
      Caption         =   "Dimes"
      Height          =   253
      Index           =   1
      Left            =   10
      TabIndex        =   3
      Top             =   1177
      Width           =   605
   End
   Begin VB.Label Coin 
      Alignment       =   1  'Right Justify
      Caption         =   "Quarters"
      Height          =   253
      Index           =   0
      Left            =   25
      TabIndex        =   0
      Top             =   935
      Width           =   594
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'Editorial Comment: I've always felt that Visual Basic is preferable
' to other programming languages such as 'C' because Visual Basic programs
' are easier to read and understand. This example should convince any
' skeptics of the validity of this opinion. Not only is the code in this
' example self-documenting, the power and flexibility of Visual Basic allows
' the programmer to create compact and elegant solutions.
'
'This program solves a homework problem with fewest possible lines of code:
' "For value in 'balance' textbox, calculate minimum number of
' quarters, dimes, nickels, and pennies required to make up given total.
' Display results in labeled textboxes."
'
' There is no requirement that the application be usable more than once,
' so there is no code to clear / reset anything - it will work only once,
' then must be restarted.
'
' All variables must be declared explicitly - that's why there are none.
'
' This one event subroutine is all the code there is. In addition to this
' code, there is exactly one form with five text boxes (balance and five
' boxes for coin counts) and six labels ('Balance', coin names, and
' instructions).
'
' There is nothing else.
'
' No effort is wasted to make this solution comprehensible.
' Since there's only seven lines of code, and only three lines that
' perform computation, how bad could it be?
'
Private Sub Balance_lostfocus()
' This event routine is invoked when the user exits the 'balance' box.
  
  If Balance >= Coin(Balance.Left).Left Then
    Coins(Balance.Left) = Coins(Balance.Left) + 1
    ' Keep decrementing until there's no balance left
    Balance = Balance - Coin(Balance.Left).Left
  Else
    Balance.Left = Balance.Left + 1
  End If
    
  If Balance Then Balance_lostfocus

End Sub

