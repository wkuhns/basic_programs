Option Explicit

Dim buff As String
Dim fullbuff As Integer

Function getcom () As String

    ComForm!Timer1.Enabled = False
    If fullbuff Then
        getcom = Left$(buff, InStr(buff, Chr$(13)) - 1)
        buff = Right$(buff, Len(buff) - InStr(buff, Chr$(10)))
        ComForm!Text2.Text = buff
        fullbuff = False
    Else
        getcom = ""
    End If
    ComForm!Timer1.Enabled = True

End Function

Sub LoadCommBuff ()

    buff = buff + ComForm!Comm1.Input
    ComForm!Text2.Text = buff
    If InStr(buff, Chr$(10)) Then
        fullbuff = True
    End If

End Sub

Sub OpenCom ()

    ComForm!Comm1.CommPort = 2
    ComForm!Comm1.InputLen = 0
    ComForm!Comm1.PortOpen = True
    fullbuff = False

End Sub

Sub PsuedoGet ()

    buff = buff + ComForm!Text4.Text
    fullbuff = True
End Sub

Sub SendCom (buffer As String)

    Dim mybuff As String

    ComForm!Comm1.InputLen = 0

    ComForm!Comm1.Output = buffer + Chr$(13) + Chr$(10)
'    While ComForm!Comm1.InBufferCount = 0
'        DoEvents
'    Wend
'    mybuff = ComForm!Comm1.Input
    ComForm!Text1.Text = buffer

End Sub

