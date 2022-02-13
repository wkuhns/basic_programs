VERSION 5.00
Begin VB.Form CCCForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "CCC"
   ClientHeight    =   3060
   ClientLeft      =   105
   ClientTop       =   7110
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3060
   ScaleWidth      =   5175
   Begin VB.CommandButton ClearButton 
      Appearance      =   0  'Flat
      Caption         =   "Clear"
      Height          =   372
      Left            =   4200
      TabIndex        =   1
      Top             =   2640
      Width           =   972
   End
   Begin VB.ListBox CCCBox 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5172
   End
End
Attribute VB_Name = "CCCForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClearButton_Click()

    CCCBox.Clear

End Sub

