VERSION 5.00
Begin VB.Form HOFForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Hall Of Fame"
   ClientHeight    =   2390
   ClientLeft      =   880
   ClientTop       =   1520
   ClientWidth     =   4270
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2390
   ScaleWidth      =   4270
   Begin VB.ListBox HOFBox 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4332
   End
End
Attribute VB_Name = "HOFForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

