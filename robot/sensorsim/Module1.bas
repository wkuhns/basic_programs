Attribute VB_Name = "Module1"
Type coord
   x As Integer
   y As Integer
End Type

Type botstruct
    x As Integer
    y As Integer
    t As Integer
End Type

Public bot As botstruct

Public xscale As Single
Public yscale As Single
Public yoffset As Single

Public scan(36) As Integer

Public runflag As Boolean

