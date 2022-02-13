Attribute VB_Name = "Module1"
Option Explicit

Global Const numcolors = 6
Global Const numpegs = 4

Global patterns(numcolors, numcolors, numcolors, numcolors) As String * 1
Global hidden(numpegs) As Integer
Global Chidden(numpegs) As Integer
Global guess(numpegs) As Integer
Global Turn As Integer

Global ChosenColor As Integer
Global WhiteCount As Integer
Global Blackcount As Integer

Global Const BLACK = 0
Global Const WHITE = &HFFFFFF



Sub ClearPattern()

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer

For i = 1 To numcolors
    For j = 1 To numcolors
        For k = 1 To numcolors
            For l = 1 To numcolors
                patterns(i, j, k, l) = "P"
            Next l
        Next k
    Next j
Next i

End Sub


Sub score(h() As Integer, g() As Integer)

    Dim i As Integer
    Dim j As Integer
    Dim index As Integer
    
    Dim hiddenx(numpegs) As Integer
    Dim guessx(numpegs) As Integer
    
    ' make a copy
    For i = 0 To numpegs - 1
        guessx(i) = g(i)
        hiddenx(i) = h(i)
    Next i
    
    ' count blacks
    Blackcount = 0
    For i = 0 To numpegs - 1
        If hiddenx(i) = guessx(i) Then
            Blackcount = Blackcount + 1
            hiddenx(i) = -1
            guessx(i) = -2
        End If
    Next i
    
    ' count whites
    WhiteCount = 0
    For i = 0 To numpegs - 1
        For j = 0 To numpegs - 1
           If hiddenx(i) = guessx(j) Then
                WhiteCount = WhiteCount + 1
                hiddenx(i) = -1
                guessx(j) = -2
           End If
        Next j
    Next i
    
    
End Sub

Function GetIndex(t As Integer, i As Integer)

GetIndex = t * numpegs + i

End Function

