VERSION 2.00
Begin Form Form3 
   Caption         =   "Display"
   ClientHeight    =   1728
   ClientLeft      =   7764
   ClientTop       =   4512
   ClientWidth     =   1692
   Height          =   2148
   Left            =   7716
   LinkTopic       =   "Form3"
   ScaleHeight     =   1728
   ScaleWidth      =   1692
   Top             =   4140
   Width           =   1788
   Begin HScrollBar ContrastBar 
      Height          =   255
      Left            =   240
      Max             =   127
      Min             =   64
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Value           =   96
      Width           =   1215
   End
   Begin CheckBox BackLightBox 
      Caption         =   "BackLight"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   1215
   End
   Begin Label Label1 
      Caption         =   "Contrast"
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   9.6
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End

Sub BackLightBox_Click ()

    If BackLightBox.Value Then
        BacklightOn
    Else
        BacklightOff
    End If

End Sub

Sub ContrastBar_Change ()
    
    SetContrast (ContrastBar.Value)
    
End Sub

