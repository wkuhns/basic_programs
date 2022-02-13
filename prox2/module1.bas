Type message
    vectors(2, 6) As Integer
    TopLine As String
    BotLine As String
End Type

Type menustate
    choices As Integer
    msg(20) As message
    choice As Integer
End Type

Global Menus(12) As menustate

Global SystemState As Integer
Global SystemLine As Integer

