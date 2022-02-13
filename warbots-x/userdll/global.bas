Attribute VB_Name = "Global"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public MyForm As StatForm
Public Myrobot As ServerRobot

Public stat As String
Public myx As Integer
Public myy As Integer
Public myheat As Integer
Public myhealth As Integer
Public mydirection As Integer
Public myspeed As Integer
Public mytime As Integer

Public lastaccess As Long



