VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MapForm 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Map"
   ClientHeight    =   8520
   ClientLeft      =   4068
   ClientTop       =   1008
   ClientWidth     =   9240
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   Picture         =   "map.frx":0000
   ScaleHeight     =   8520
   ScaleWidth      =   9240
   Begin VB.TextBox DispBox 
      Height          =   372
      Index           =   6
      Left            =   6840
      TabIndex        =   57
      Top             =   7800
      Width           =   372
   End
   Begin VB.TextBox DispBox 
      Height          =   372
      Index           =   5
      Left            =   6480
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   7800
      Width           =   372
   End
   Begin VB.TextBox DispBox 
      Height          =   372
      Index           =   4
      Left            =   6120
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   7800
      Width           =   372
   End
   Begin VB.TextBox DispBox 
      Height          =   372
      Index           =   3
      Left            =   5760
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   7800
      Width           =   372
   End
   Begin VB.TextBox DispBox 
      Height          =   372
      Index           =   2
      Left            =   5400
      TabIndex        =   53
      Text            =   "Text1"
      Top             =   7800
      Width           =   372
   End
   Begin VB.TextBox DispBox 
      Height          =   372
      Index           =   1
      Left            =   5040
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   7800
      Width           =   372
   End
   Begin VB.TextBox DispBox 
      Height          =   372
      Index           =   0
      Left            =   4680
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   7800
      Width           =   372
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   4560
      Top             =   0
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   713
   End
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   5520
      Top             =   0
      _ExtentX        =   635
      _ExtentY        =   635
      _Version        =   393216
   End
   Begin VB.CommandButton SightButton 
      Appearance      =   0  'Flat
      Caption         =   "S+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1200
      TabIndex        =   50
      Top             =   7800
      Width           =   372
   End
   Begin VB.ComboBox SightList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   6240
      TabIndex        =   49
      Top             =   0
      Width           =   3012
   End
   Begin VB.CommandButton EndButton 
      Appearance      =   0  'Flat
      Caption         =   "End Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   48
      Top             =   7800
      Width           =   972
   End
   Begin VB.CommandButton AllButton 
      Appearance      =   0  'Flat
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   600
      TabIndex        =   47
      Top             =   7800
      Width           =   612
   End
   Begin VB.TextBox CoordBox 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2880
      TabIndex        =   0
      Top             =   0
      Width           =   1212
   End
   Begin VB.CommandButton ClearButton 
      Appearance      =   0  'Flat
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   0
      TabIndex        =   6
      Top             =   7800
      Width           =   612
   End
   Begin VB.TextBox ClockBox 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   960
      TabIndex        =   5
      Top             =   0
      Width           =   972
   End
   Begin VB.TextBox MessageBox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   0
      TabIndex        =   4
      Top             =   8160
      Width           =   7812
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   0
   End
   Begin VB.CommandButton RunButton 
      Appearance      =   0  'Flat
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1920
      TabIndex        =   3
      Top             =   0
      Width           =   972
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   7452
      Left            =   9000
      Picture         =   "map.frx":0446
      ScaleHeight     =   7428
      ScaleWidth      =   228
      TabIndex        =   2
      Top             =   360
      Width           =   252
   End
   Begin VB.CommandButton PauseButton 
      Appearance      =   0  'Flat
      Caption         =   "Pause"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   972
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   0
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   1
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   2
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   3
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   4
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   5
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   6
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   7
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   8
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   9
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   10
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   11
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   12
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   13
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   14
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   15
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   16
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   17
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   18
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   19
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   20
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   21
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   22
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   23
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   24
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   25
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   26
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   27
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   28
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   29
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   30
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   31
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   32
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   33
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   34
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   35
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   36
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   37
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   38
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   45
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox DItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Index           =   39
      Left            =   4080
      ScaleHeight     =   96
      ScaleWidth      =   96
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Menu topmenu 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu FileMenu 
         Caption         =   "&New"
         Index           =   0
      End
      Begin VB.Menu FileMenu 
         Caption         =   "&Load Game"
         Index           =   1
      End
      Begin VB.Menu FileMenu 
         Caption         =   "&Save Game"
         Index           =   2
      End
      Begin VB.Menu FileMenu 
         Caption         =   "&Print"
         Index           =   3
      End
      Begin VB.Menu FileMenu 
         Caption         =   "&Quit"
         Index           =   4
      End
   End
   Begin VB.Menu topmenu 
      Caption         =   "&Display"
      Index           =   1
      Begin VB.Menu DisplayMenu 
         Caption         =   "Redraw Map"
         Index           =   0
      End
      Begin VB.Menu DisplayMenu 
         Caption         =   "Show Difficulty"
         Index           =   1
      End
      Begin VB.Menu DisplayMenu 
         Caption         =   "Show Cover"
         Index           =   2
      End
      Begin VB.Menu DisplayMenu 
         Caption         =   "Show Rivers"
         Index           =   3
      End
      Begin VB.Menu DisplayMenu 
         Caption         =   "Show Water"
         Index           =   4
      End
   End
   Begin VB.Menu topmenu 
      Caption         =   "&List"
      Index           =   2
      Begin VB.Menu ListMenu 
         Caption         =   "Moving Units"
         Index           =   0
      End
      Begin VB.Menu ListMenu 
         Caption         =   "All Units"
         Index           =   1
      End
      Begin VB.Menu ListMenu 
         Caption         =   "Stopped Units"
         Index           =   2
      End
   End
   Begin VB.Menu topmenu 
      Caption         =   "&Buy"
      Index           =   3
   End
   Begin VB.Menu topmenu 
      Caption         =   "&Rain"
      Index           =   4
      Begin VB.Menu RainMenu 
         Caption         =   "Rain"
         Index           =   0
      End
      Begin VB.Menu RainMenu 
         Caption         =   "Flow"
         Index           =   1
      End
      Begin VB.Menu RainMenu 
         Caption         =   "Flood"
         Index           =   2
      End
   End
   Begin VB.Menu topmenu 
      Caption         =   "&Windows"
      Index           =   5
      Begin VB.Menu WindowMenu 
         Caption         =   "Colors"
         Index           =   0
      End
      Begin VB.Menu WindowMenu 
         Caption         =   "Sightings"
         Index           =   1
      End
      Begin VB.Menu WindowMenu 
         Caption         =   "CCC"
         Index           =   2
      End
   End
   Begin VB.Menu topmenu 
      Caption         =   "Mouseflavor"
      Index           =   6
      Begin VB.Menu MouseMenu 
         Caption         =   "Coordinates"
         Index           =   0
      End
      Begin VB.Menu MouseMenu 
         Caption         =   "Rain"
         Index           =   1
      End
      Begin VB.Menu MouseMenu 
         Caption         =   "Difficulty"
         Index           =   2
      End
      Begin VB.Menu MouseMenu 
         Caption         =   "Altitude"
         Index           =   3
      End
      Begin VB.Menu MouseMenu 
         Caption         =   "Cover"
         Index           =   5
      End
   End
End
Attribute VB_Name = "MapForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public mouseflavor As Integer
Dim OldX As Single
Dim OldY As Single
Dim StartX As Single
Dim StartY As Single
'* Mouse locations

Public Sub DrawSelectionBox(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
    Dim dmode As Integer
    
    dmode = Me.DrawMode
Me.DrawMode = INVERSE
Me.DrawWidth = 3
Me.FillStyle = 1
Me.Line (x1 * mapdx + mapleft, y1 * mapdy + maptop)-(x2 * mapdx + mapleft, y2 * mapdy + maptop), , B

Me.DrawMode = dmode

End Sub
Public Sub DrawBox(x1 As Single, y1 As Single, x2 As Single, y2 As Single)
    Dim dmode As Integer
    
    dmode = Me.DrawMode
Me.DrawMode = INVERSE
Me.DrawWidth = 3
Me.FillStyle = 1
Me.Line (x1, y1)-(x2, y2), , B

Me.DrawMode = dmode

End Sub

' Process user's choice from 'Display' menu
Private Sub DisplayMenu_Click(index As Integer)

    Select Case index
        Case 0
            DisplayMap
        Case Else
            DisplayMap2 (index)
    End Select
        
End Sub


Private Sub Form_Unload(Cancel As Integer)

    If Winsock.State <> sckClosed Then Winsock.Close

End Sub

Private Sub MouseMenu_Click(index As Integer)

    mouseflavor = index
    
End Sub

Private Sub RainMenu_Click(index As Integer)

Select Case index
    Case 0
        MakeRain
    Case 1
        MakeFlow
    Case 2
        MakeFlood
End Select

End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim strdata As String

    On Error GoTo finis
    Winsock.GetData strdata, vbString
    combuff = combuff + strdata
    ParseCommand
    Exit Sub
finis:
    DoEvents
    Winsock.GetData strdata, vbString
    combuff = combuff + strdata
    If combuff <> "" Then
        ParseCommand
    End If
    
End Sub

' Show all our units
Private Sub AllButton_Click()

    Dim i As Integer

    If side = us Then
        For i = 0 To asize - 1
            If army(0, i).health > 0 Then
                AddDisplayUnit i, DSQUARE, BLACK, WHITE
            End If
        Next i
    Else
        For i = 0 To asize - 1
            If army(1, i).health > 0 Then
                AddDisplayUnit i, DSQUARE, BLACK, WHITE
            End If
        Next i
    End If

    PlotDisplayItems

End Sub

Private Sub ClearButton_Click()

    Dim i As Integer

    For i = 0 To 39
        MapForm!DItem(i).Visible = False
    Next i

    MapForm!SightList.Clear
    dlcount = 0

End Sub

Private Sub DispBox_Click(index As Integer)
    
    Dim i As Integer

'    SavePicture DispBox(0).Image, "Box.bmp"

    If side = us Then
        For i = 0 To asize - 1
            If (army(0, i).health > 0) And (army(0, i).type = index) Then
                AddDisplayUnit i, DSQUARE, BLACK, WHITE
            End If
        Next i
    Else
        For i = 0 To asize - 1
            If (army(1, i).health > 0) And (army(1, i).type = index) Then
                AddDisplayUnit i, DSQUARE, BLACK, WHITE
            End If
        Next i
    End If

    PlotDisplayItems

End Sub

Private Sub DispBox_DblClick(index As Integer)

    MakePickList index

End Sub

Private Sub DItem_Click(index As Integer)

    Dim i As Integer
    Dim lcount As Integer
    Dim unit As Integer

    If UnitWaiting Then
        Form_MouseDown 1, 0, (DItem(index).Left + (DItem(index).Width / 2)), (DItem(index).Top + DItem(index).Height / 2)
    Else
        CommandForm!UnitBox.Clear
        lcount = 0
        If side = us Then
            For i = 0 To asize - 1
                If (Int(army(0, i).X) = dc(index).X) And (Int(army(0, i).Y) = dc(index).Y) Then
                    AddListItem army(0, i)
                    lcount = lcount + 1
                    unit = i
                End If
            Next i
        Else
            For i = 0 To asize - 1
                If Int(army(1, i).X) = dc(index).X And (Int(army(1, i).Y) = dc(index).Y) Then
                    AddListItem army(1, i)
                    lcount = lcount + 1
                    unit = i
                End If
            Next i
        End If

        If CommandForm!UnitBox.ListCount > 0 Then
            CommandForm!UnitBox.Selected(0) = True
            CommandForm.Show
        End If

    End If

End Sub

Private Sub DrawMapForm()

    Dim i As Integer

    MapForm.FontSize = 8

    maptop = PauseButton.Height + 1.5 * TextHeight("X")
    mapleft = 1.5 * TextWidth("XX")
    MapHeight = ClearButton.Top - maptop
    MapWidth = (MapForm.Width - mapleft) * 0.95
    mapdy = MapHeight / axis
    mapdx = MapWidth / axis

    Picture2.Top = maptop
    Picture2.Width = (MapForm.Width - mapleft) * 0.04
    Picture2.Left = mapleft + (MapForm.Width - mapleft) * 0.96
    Picture2.Height = MapHeight

    For i = 0 To 39
        DItem(i).Visible = False
        DItem(i).FillStyle = 0
        DItem(i).Height = 2 * mapdy
        DItem(i).Width = 2 * mapdx
    Next i

    MapForm.Show
    MapForm.AutoRedraw = True

    For i = 0 To 9
        MapForm.CurrentY = PauseButton.Height
        MapForm.CurrentX = (MapWidth / 10) * i + mapleft
        MapForm.Print Str$(i * 10)
        MapForm.CurrentY = maptop + MapHeight / 10 * i
        MapForm.CurrentX = 0
        MapForm.Print Str$(i * 10)
    Next i

End Sub

Private Sub EndButton_Click()

    UnitWaiting = False
    CommandForm!UnitBox.Enabled = True
    MapForm.MousePointer = 0

End Sub

Private Sub FileMenu_Click(index As Integer)

    Select Case index
        Case 0                  ' new
            InitMap
            DisplayMap
            SetupPlayer
            CCCForm.Show
        Case 1                  'Load
            LoadGame
            DisplayMap
            CCCForm.Show
        Case 2                  'save
            SaveGame
        Case 3                  ' print
            PrintMap
        Case 4                  ' quit
            End
    End Select

End Sub

Private Sub Form_Load()
    
    Dim i As Integer

    If Winsock.State <> sckClosed Then Winsock.Close

    If Winsock.LocalHostName = "wiley" Then
        Winsock.RemoteHost = "demon"
        Winsock.RemotePort = 1001
        Winsock.Bind 1002
    Else
        Winsock.RemoteHost = "wiley"
        Winsock.RemotePort = 1002
        Winsock.Bind 1001
    End If
    
    MessageBox.Text = Winsock.LocalHostName
        
    InitGame                                            ' set colors, vectors
    LoadUnits           ' get unit specs

    DrawMapForm

    For i = 0 To 39
        DItem(i).FontSize = 5
    Next i
    MessageBox.Text = MessageBox.Text & "...done"
    remote = 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Select Case Button
        Case 1
            If UnitWaiting Then
                DispatchUnits X, Y
            End If
            '* Store the initial start of the line to draw.
            StartX = X
            StartY = Y
        
            '* Make the last location equal the starting location
            OldX = StartX
            OldY = StartY
            
            LastClick.X = (X - mapleft) / mapdx
            LastClick.Y = (Y - maptop) / mapdy

    End Select

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim x1 As Integer
    Dim y1 As Integer
    
    If Button = 1 Then
       '* Erase the previous line.
       Call DrawBox(StartX, StartY, OldX, OldY)
    
       '* Draw the new line.
       Call DrawBox(StartX, StartY, X, Y)
    
       '* Save the coordinates for the next call.
       OldX = X
       OldY = Y
       Exit Sub
    End If
    
    Select Case mouseflavor
        Case 0
            CoordBox.Text = Format$(Int((X - mapleft) / mapdx), "00") + ", " + Format$(Int((Y - maptop) / mapdy), "00")
        Case 1
            x1 = Int((X - mapleft) / mapdx)
            y1 = Int((Y - maptop) / mapdy)
            If x1 > 0 And x1 < 100 And y1 > 0 And y1 < 100 Then
                CoordBox.Text = rain(x1, y1)
            End If
        Case 2
            x1 = Int((X - mapleft) / mapdx)
            y1 = Int((Y - maptop) / mapdy)
            If x1 > 0 And x1 < 100 And y1 > 0 And y1 < 100 Then
                CoordBox.Text = terrain(x1, y1).d
            End If
        Case 3
            x1 = Int((X - mapleft) / mapdx)
            y1 = Int((Y - maptop) / mapdy)
            If x1 > 0 And x1 < 100 And y1 > 0 And y1 < 100 Then
                CoordBox.Text = terrain(x1, y1).a
            End If
        Case 4
            x1 = Int((X - mapleft) / mapdx)
            y1 = Int((Y - maptop) / mapdy)
            If x1 > 0 And x1 < 100 And y1 > 0 And y1 < 100 Then
                CoordBox.Text = terrain(x1, y1).t
            End If
    End Select

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Abs(X - StartX) > mapdx * 5 And Abs(Y - StartY) > mapdy * 5 Then
        Call DrawBox(StartX, StartY, OldX, OldY)
        SubMap.SetSubMap 1, (StartX - mapleft) / mapdx, (StartY - maptop) / mapdy, (X - mapleft) / mapdx, (Y - maptop) / mapdy
    End If
End Sub

Private Sub ListMenu_Click(index As Integer)

    Dim i As Integer

    CommandForm!UnitBox.Clear
    
    Select Case index
        Case 0
            If side = us Then
                For i = 0 To asize - 1
                    If (army(0, i).health > 0) And (army(0, i).speed > 0) Then
                        AddListItem army(0, i)
                    End If
                Next i
            Else
                For i = 0 To asize - 1
                    If (army(1, i).health > 0) And (army(1, i).speed > 0) Then
                        AddListItem army(1, i)
                    End If
                Next i
            End If
        Case 1
            If side = us Then
                For i = 0 To asize - 1
                    If army(0, i).health > 0 Then
                        AddListItem army(0, i)
                    End If
                Next i
            Else
                For i = 0 To asize - 1
                    If army(1, i).health > 0 Then
                        AddListItem army(1, i)
                    End If
                Next i
            End If
        Case 2
            If side = us Then
                For i = 0 To asize - 1
                    If (army(0, i).health > 0) And (army(0, i).speed = 0) Then
                        AddListItem army(0, i)
                    End If
                Next i
            Else
                For i = 0 To asize - 1
                    If (army(1, i).health > 0) And (army(1, i).speed = 0) Then
                        AddListItem army(1, i)
                    End If
                Next i
            End If
    End Select
    CommandForm.Show

End Sub

Private Sub PauseButton_Click()

    Timer1.Enabled = False
    PauseButton.Enabled = False
    RunButton.Enabled = True
    SendCom "P"

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i, dh As Integer

    FileDialog.Action = 3           ' color dialog
    dh = Int(Picture2.Height / 10)
    i = Int(Y / dh)

    mc(9 - i) = FileDialog.Color
    Picture2.Line (0, i * dh)-Step(Picture2.Width, dh), mc(9 - i), BF

End Sub

Public Sub RunButton_Click()

    Timer1.Enabled = True
    PauseButton.Enabled = True
    RunButton.Enabled = False
    SendCom "R"
End Sub


Private Sub SightButton_Click()

    AddSightings
    PlotDisplayItems

End Sub

Private Sub Timer1_Timer()
    tick
End Sub

Private Sub TopMenu_Click(index As Integer)

    Dim i As Integer

    Select Case index
        Case 3
            BuyForm.Show

    End Select

End Sub


Private Sub WindowMenu_Click(index As Integer)

    Select Case index
        Case 0              ' color
'            ColorForm.Show
            GetColors
        Case 1              ' Sightings
            Call SubMap.SetSubMap(1, 10, 10, 30, 30)
        Case 2             ' CCC
            CCCForm.Show
    End Select

End Sub

