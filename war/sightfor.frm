VERSION 2.00
Begin Form SightForm 
   Caption         =   "Sightings"
   ClientHeight    =   3336
   ClientLeft      =   8556
   ClientTop       =   600
   ClientWidth     =   2892
   Height          =   3756
   Left            =   8508
   LinkTopic       =   "Form1"
   ScaleHeight     =   3336
   ScaleWidth      =   2892
   Top             =   228
   Width           =   2988
   Begin CommandButton SightButton 
      Caption         =   "Clear"
      Height          =   492
      Index           =   3
      Left            =   2160
      TabIndex        =   4
      Top             =   2880
      Width           =   732
   End
   Begin CommandButton SightButton 
      Caption         =   "ReDraw"
      Height          =   492
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Top             =   2880
      Width           =   732
   End
   Begin CommandButton SightButton 
      Caption         =   "Delete"
      Height          =   492
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   2880
      Width           =   732
   End
   Begin CommandButton SightButton 
      Caption         =   "HiLite"
      Height          =   492
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   2880
      Width           =   732
   End
   Begin ListBox SightList 
      Height          =   2904
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   0
      Width           =   2892
   End
End
Option Explicit

Sub SightButton_Click (index As Integer)

    Select Case index
        Case 0                  ' HiLite
            ' (draw highlight on map)
        Case 1                  ' Delete
            sightList.RemoveItem sightList.ListIndex

'            RemoveDisplayItems "S"
'            AddSightings
        Case 2                  ' Redraw
            AddSightings
            PlotDisplayItems
        Case 3                  ' Clear
            sightList.Clear
'            RemoveDisplayItems "S"
    End Select

End Sub

