Attribute VB_Name = "robotmod"
Option Explicit
' We might need a pointer to the common server object...
Public grServer As rServer
Public gUseCount As Integer

' To-do list:
' - move dead 'bots off field
' - add redraw command to api
' - enhance mark command
' - tournament play
' - check cooldown logic / rate
' - better boom graphics

