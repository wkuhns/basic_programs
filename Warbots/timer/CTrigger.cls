VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTrigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function SetTimerAPI Lib "user32" Alias "SetTimer" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimerAPI Lib "user32" Alias "KillTimer" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private WithEvents m_TimerSink As CTimerSink
Attribute m_TimerSink.VB_VarHelpID = -1
Private m_nID As Long
Private m_Interval As Long
Private m_Tag As Variant

Public Event Trigger()

Public Property Get Tag() As Variant
    If Not IsEmpty(m_Tag) Then
        If IsObject(m_Tag) Then
            Set Tag = m_Tag
        Else
            Tag = m_Tag
        End If
    End If
End Property

Public Property Let Tag(vTag As Variant)
    m_Tag = vTag
End Property

Public Property Set Tag(vTag As Variant)
    Set m_Tag = vTag
End Property

Public Sub Start(Optional ByVal iInterval As Long = 100&)
    m_Interval = iInterval
    
    StartTimer
End Sub

Private Sub StartTimer()
    If m_nID <> 0 Then
        StopTimer
    End If
    
    If m_Interval > 0 Then
        'The return from SetTimerAPI is a unique handle to this timer, that
        'we'll use later to decide whether a CTimerSink event is directed at
        'this object
        m_nID = SetTimerAPI(0&, 0&, m_Interval, AddressOf modTimer.TimerProc)
    End If
End Sub

Private Sub StopTimer()
    If m_nID <> 0 Then
        KillTimerAPI 0&, m_nID
        m_nID = 0
    End If
End Sub

Private Sub Class_Initialize()
    Set m_TimerSink = modTimer.TimerSink
    m_Interval = 100
End Sub

Private Sub Class_Terminate()
    StopTimer
    
    Set m_TimerSink = Nothing
End Sub

Private Sub m_TimerSink_Timer(ByVal nID As Long)
    If nID = m_nID Then
        'This timer event is meant for this object so kill the timer
        'and raise an event.
        StopTimer
        RaiseEvent Trigger
    End If
End Sub

