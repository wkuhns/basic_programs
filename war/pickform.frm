VERSION 2.00
Begin Form PickForm 
   Caption         =   "Units"
   ClientHeight    =   2940
   ClientLeft      =   4620
   ClientTop       =   4764
   ClientWidth     =   2880
   Height          =   3360
   Left            =   4572
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   2880
   Top             =   4392
   Width           =   2976
   Begin CommandButton EndButton 
      Caption         =   "End"
      Height          =   372
      Left            =   1440
      TabIndex        =   4
      Top             =   2520
      Width           =   612
   End
   Begin CommandButton PathButton 
      Caption         =   "Path"
      Height          =   372
      Left            =   720
      TabIndex        =   3
      Top             =   2520
      Width           =   612
   End
   Begin CommandButton OrderButton 
      Caption         =   "Orders"
      Height          =   372
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   612
   End
   Begin CommandButton Command1 
      Caption         =   "Done"
      Height          =   372
      Left            =   2160
      TabIndex        =   1
      Top             =   2520
      Width           =   612
   End
   Begin ListBox PickList 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Courier"
      FontSize        =   9.6
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1944
      Left            =   0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   2892
   End
End
Option Explicit

Sub Command1_Click ()
    
    PickList.Clear
    PickForm.Hide

End Sub

Sub EndButton_Click ()
    
    UnitWaiting = False
    PickList.Enabled = True
    MapForm.MousePointer = 0

End Sub

Sub PathButton_Click ()

    UnitWaiting = True
    PickList.Enabled = False
    MapForm.MousePointer = 2

End Sub

Sub PickList_DblClick ()
            
    Dim i As Integer

    i = Val(Left$(PickList.Text, 2))
    
    If side = US Then
        MakeCommandForm a1(i)
    Else
        MakeCommandForm a2(i)
    End If

End Sub

