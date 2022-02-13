Attribute VB_Name = "modTimer"
Option Explicit
   
Private m_TimerSink As CTimerSink

Public Property Get TimerSink() As CTimerSink
    If m_TimerSink Is Nothing Then
        Set m_TimerSink = New CTimerSink
    End If
    
    Set TimerSink = m_TimerSink
End Property
    
Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal nID As Long, ByVal dwTime As Long)
    'Raise a public event in each object in this component and let them decide who
    'gets it.
    TimerSink.Timer nID
End Sub
